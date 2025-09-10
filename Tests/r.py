import logging
import os
import uuid
from typing import Optional
from typing import Union
import subprocess
import tiktoken
import asyncio
from asyncio import Semaphore
from pypdf import PdfReader
import tempfile
import shutil

import openai
from openai.types.file_object import FileObject

from open_webui.apps.webui.models.files import FileForm, FileModel, Files
from open_webui.config import OPENAI_API_KEY, UPLOAD_SEMAPHORE_LIMIT, ENABLE_PDF_CONVERSION, CONVERSION_FORMAT, TOKEN_LIMIT, KNOWLEDGE_SUPPORTED_EXTENSIONS
from open_webui.env import SRC_LOG_LEVELS
from open_webui.constants import ERROR_MESSAGES
from open_webui.apps.webui.routers.files import upload_file, get_file_by_id

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile, status
from fastapi.responses import JSONResponse, StreamingResponse

from open_webui.utils.utils import get_admin_user, get_verified_user

log = logging.getLogger(_name_)
log.setLevel(SRC_LOG_LEVELS["MODELS"])

router = APIRouter()

# Configure OpenAI
openai.api_key = OPENAI_API_KEY

# Create a semaphore to limit concurrent jobs
upload_semaphore = Semaphore(UPLOAD_SEMAPHORE_LIMIT)

############################
# Upload File to OpenAI
############################

# Move this to the module level (outside any function)
async def extract_pdf_text(pdf_path: str) -> str:
    """
    Extract text from a PDF file.
    
    Args:
        pdf_path: Path to the PDF file
        
    Returns:
        Extracted text from all pages
    """
    def _process_pdf(path):
        try:
            with open(path, "rb") as f:
                pdf = PdfReader(f)
                text = ""
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text
                return text
        except Exception as pypdf_error:
            log.error(f"Failed to process PDF {path}: {pypdf_error}")
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Failed to process PDF file: {str(pypdf_error)}"
            )

    # Run CPU-bound task in a thread pool
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _process_pdf, pdf_path)

# Move this to the module level (outside any function)
async def upload_file_to_openai(file_path: str, purpose: str = "user_data"):
    """
    Asynchronously upload a file to OpenAI.
    
    Args:
        file_path: Path to the file to upload
        purpose: Purpose of the file (default: "user_data")
        
    Returns:
        OpenAI API response
    """
    def _run_openai_upload(path):
        with open(path, "rb") as f:
            return openai.files.create(
                file=(os.path.basename(path), f),
                purpose=purpose
            )

    # Run the synchronous OpenAI API call in a thread pool
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _run_openai_upload, file_path)

async def convert_to_pdf_async(original_path: str, file_ext: str) -> str:
    """
    Convert a file to PDF format if needed.
    
    Args:
        original_path: Path to the original file
        file_ext: File extension (lowercase)
        
    Returns:
        Path to the PDF file (either converted or original)
        
    Raises:
        HTTPException: If conversion is required but fails, or if file type is unsupported
    """
    # https://platform.openai.com/docs/guides/tools-file-search
    # These are the extensions that OpenAI supports for tools-file-search but still requires a PDF conversion, 
    # so we need to handle it.

    # Dynamic extensions based on feature flag
    BASE_EXTENSIONS = ['.pdf']
    CONVERTIBLE_EXTENSIONS = [
        '.doc', '.docx', '.odt', '.ppt', '.pptx', '.txt',
        '.xls', '.csv', '.xlsx', '.py', '.cpp', '.c', '.md',
        '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp'
    ] if ENABLE_PDF_CONVERSION else []

    SUPPORTED_EXTENSIONS = BASE_EXTENSIONS + CONVERTIBLE_EXTENSIONS

    if file_ext not in SUPPORTED_EXTENSIONS:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unsupported file type: {file_ext}. Supported types are: {', '.join(SUPPORTED_EXTENSIONS)}"
        )

    # Conversion logic
    if ENABLE_PDF_CONVERSION and file_ext in CONVERTIBLE_EXTENSIONS:
        pdf_path = os.path.splitext(original_path)[0] + f".{CONVERSION_FORMAT}"

        # Create a temporary directory for this conversion to avoid LibreOffice profile conflicts
        temp_profile_dir = tempfile.mkdtemp(prefix="libreoffice_profile_")

        try:
            # Run LibreOffice conversion with a separate user profile to avoid conflicts
            process = await asyncio.create_subprocess_exec(
                "libreoffice",
                f"-env:UserInstallation=file://{temp_profile_dir}",
                "--headless",
                "--convert-to", CONVERSION_FORMAT,
                "--outdir", os.path.dirname(original_path),
                original_path,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE
            )

            stdout, stderr = await process.communicate()

            if process.returncode != 0:
                error_msg = stderr.decode().strip() or "Unknown conversion error"
                log.error(f"Conversion failed for {original_path}: {error_msg}")
                log.error(f"LibreOffice stdout: {stdout.decode().strip()}")
                raise HTTPException(
                    status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                    detail="Failed to convert file to PDF"
                )
        finally:
            # Clean up the temporary profile directory
            try:
                shutil.rmtree(temp_profile_dir)
            except Exception as cleanup_error:
                log.warning(f"Failed to cleanup temp profile directory {temp_profile_dir}: {cleanup_error}")

        return pdf_path
    else:
        if file_ext != '.pdf':
            raise HTTPException(
                status_code=400,
                detail=f"Direct PDF upload required for {file_ext} files"
            )
        return original_path

async def validate_token_count_async(pdf_path: str) -> None:
    """
    Validate that the PDF file doesn't exceed the token limit.
    
    Args:
        pdf_path: Path to the PDF file
        
    Raises:
        HTTPException: If the file exceeds the token limit
    """
    # Extract text from PDF
    pdf_text = await extract_pdf_text(pdf_path)

    # Count tokens
    encoding = tiktoken.get_encoding("cl100k_base")
    tokens = encoding.encode(pdf_text)
    token_count = len(tokens)

    if token_count > TOKEN_LIMIT:  # 2M token limit
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"File exceeds 2 million token limit (contains {token_count} tokens). Please upload a file with less text."
        )

@router.post("/")
async def upload_file_openai(
    file: UploadFile = File(...), 
    user=Depends(get_verified_user),
    type: str = "chat"  # Default to chat behavior
):
    async with upload_semaphore:
        try:
            # Check if file is None, Hits only for files uploaded using google drive
            if file.size is None:
                file.size = 0
                content = await file.read()
                file.size = len(content)
                await file.seek(0)

            # Check file size limit
            if file.size > 10 * 1024 * 1024:  # 10MB
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="File size exceeds 10MB limit. Please upload a smaller file."
                )

            # Upload original file to axxessio
            uploaded_file = await upload_file(file, user)

            if type == "knowledge":
                # For knowledge uploads, upload directly to OpenAI vector store without PDF conversion
                original_path = uploaded_file.meta['path']

                # Check if file extension is supported for knowledge uploads
                file_ext = os.path.splitext(original_path)[1].lower()
                if file_ext not in KNOWLEDGE_SUPPORTED_EXTENSIONS:
                    raise HTTPException(
                        status_code=status.HTTP_400_BAD_REQUEST,
                        detail=f"Unsupported file type for knowledge: {file_ext}. Supported types are: {', '.join(KNOWLEDGE_SUPPORTED_EXTENSIONS)}"
                    )

                # Upload file directly to OpenAI using our refactored function, purpose is "assistants" and not "user_data"
                response = await upload_file_to_openai(original_path, purpose="assistants")

                log.info(f"OpenAI File ID (Vector Store): {response.id}")

                text_content = response.id

                Files.update_file_data_by_id(
                    uploaded_file.id,
                    {"content": text_content},
                )

                return [uploaded_file, response]

            else:
                # For chat uploads, convert to PDF first (existing behavior)
                original_path = uploaded_file.meta['path']
                file_ext = os.path.splitext(original_path)[1].lower()

                # Convert file to PDF if needed using the refactored function
                pdf_path = await convert_to_pdf_async(original_path, file_ext)

                # Validate token count
                await validate_token_count_async(pdf_path)

                # Upload file to OpenAI using our refactored function
                response = await upload_file_to_openai(pdf_path)

                log.info(f"OpenAI File ID: {response.id}")

                text_content = response.id

                Files.update_file_data_by_id(
                    uploaded_file.id,
                    {"content": text_content},
                )

                # Create a FileForm object to store basic info in the database.
                id = response.id
                filename = file.filename
                file_form = FileForm(
                    **{
                        "id": id,
                        "filename": filename,
                        "meta": {
                            "name": filename,
                            "content_type": file.content_type,
                            "size": file.size,
                            "openai_file_id": id,  # Store the OpenAI file ID
                        },
                    }
                )
                return [uploaded_file, response]

        except openai.APIError as e:
            log.error(f"OpenAI API Error: {e}")
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=ERROR_MESSAGES.DEFAULT(f"OpenAI API Error: {e}"),
            )
        except HTTPException as http_ex:
            # Re-raise HTTPExceptions without calling ERROR_MESSAGES.DEFAULT
            raise http_ex
        except Exception as e:
            log.exception(e)
            # Only call ERROR_MESSAGES.DEFAULT for general exceptions
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=ERROR_MESSAGES.DEFAULT(e),
            )


############################
# List Files from OpenAI
############################


@router.get("/", response_model=list[FileObject])
async def list_files_openai(user=Depends(get_verified_user)):
    try:
        # List files from OpenAI
        response = openai.files.list()

        return response.data
    except openai.APIError as e:
        log.error(f"OpenAI API Error: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=ERROR_MESSAGES.DEFAULT(f"OpenAI API Error: {e}"),
        )
    except Exception as e:
        log.exception(e)
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=ERROR_MESSAGES.DEFAULT(e),
        )


############################
# Delete All Files from OpenAI
############################


@router.delete("/all")
async def delete_all_files_openai(user=Depends(get_admin_user)):
    try:
        # List all files from OpenAI
        response = openai.files.list()
        files = response.data
        # Delete all files from OpenAI
        for file in files:
            openai.files.delete(file.id)

        return {"message": "All files deleted successfully"}
    except openai.APIError as e:
        log.error(f"OpenAI API Error: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=ERROR_MESSAGES.DEFAULT(f"OpenAI API Error: {e}"),
        )
    except Exception as e:
        log.exception(e)
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=ERROR_MESSAGES.DEFAULT(e),
        )

############################
# Delete File from OpenAI
############################


@router.delete("/{id}")
async def delete_file_openai(id: str, user=Depends(get_verified_user)):
    try:
        openai_file_id = Files.get_file_by_id(id)
        openai_file_id = openai_file_id.data['content']
        # Delete the file from OpenAI
        response = openai.files.delete(openai_file_id)
        if response.deleted:
            return {"message": "File deleted successfully"}
        else:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=ERROR_MESSAGES.DEFAULT("Error deleting file"),
            )
    except openai.APIError as e:
        log.error(f"OpenAI API Error: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=ERROR_MESSAGES.DEFAULT(f"OpenAI API Error: {e}"),
        )
    except Exception as e:
        log.exception(e)
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=ERROR_MESSAGES.DEFAULT(e),
        )


############################
# Get File By Id from OpenAI
############################


@router.get("/{id}", response_model=FileObject|FileModel)
async def get_file_by_id_openai(id: str, user=Depends(get_verified_user)):
    try:
        # Get file details from OpenAI
        response = openai.files.retrieve(id)
        return response
    except openai.NotFoundError as e:
        try:
            response = await get_file_by_id(id, user)
            return response
        except:
            log.error(f"OpenAI webui file Not Found: {e} so it must be open ai request")
        log.error(f"OpenAI File Not Found: {e}")
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=ERROR_MESSAGES.DEFAULT(f"OpenAI File Not Found: {e}"),
        )
    except openai.APIError as e:
        log.error(f"OpenAI API Error: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=ERROR_MESSAGES.DEFAULT(f"OpenAI API Error: {e}"),
        )
    except Exception as e:
        log.exception(e)
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=ERROR_MESSAGES.DEFAULT(e),
        )


############################
# Get File Content By Id from OpenAI
############################

# WR1 - This is not supported for user uploaded files.

@router.get("/{id}/content")
async def get_file_content_by_id_openai(id: str, container_id: str = None, user=Depends(get_verified_user)):
    log.info(f"Fetching file content for {id} with container_id {container_id}")
    try:
        # Check if this is a container file (from code interpreter)
        if id.startswith('cfile_'):
            # For container files, use container-specific endpoint
            # Container files require both container_id and file_id
            if not container_id:
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="Container ID is required for container files"
                )

            import httpx
            try:
                log.info(f"Fetching container file {id} with container_id {container_id}")
                async with httpx.AsyncClient() as client:
                    headers = {
                        'Authorization': f'Bearer {openai.api_key}',
                        'User-Agent': 'OpenAI/Python'
                    }

                    # For container files, try the container-specific endpoint
                    # Container files require the container ID in the URL path
                    file_id = id  # Ensure we have the correct variable name
                    container_endpoint = f'https://api.openai.com/v1/containers/{container_id}/files/{file_id}/content'
                    log.info(f"Calling OpenAI container endpoint: {container_endpoint}")

                    openai_response = await client.get(
                        container_endpoint,
                        headers=headers
                    )

                    log.info(f"OpenAI API response status: {openai_response.status_code}")

                    if openai_response.status_code != 200:
                        log.error(f"OpenAI API returned {openai_response.status_code} for container file {id}")
                        log.error(f"Response content: {openai_response.text}")
                        raise HTTPException(
                            status_code=status.HTTP_404_NOT_FOUND,
                            detail=f"Container file not found: {id} (status: {openai_response.status_code})"
                        )

                    # Get the content as bytes
                    content = openai_response.content

                    # Determine file type from content using magic bytes
                    filename = f"{id}"  # Default filename without extension
                    mime_type = "application/octet-stream"  # Default MIME type

                    # Check magic bytes to determine file type
                    if content.startswith(b'\x89PNG\r\n\x1a\n'):
                        mime_type = "image/png"
                        filename = f"{id}.png"
                    elif content.startswith(b'\xff\xd8\xff'):
                        mime_type = "image/jpeg"
                        filename = f"{id}.jpg"
                    elif content.startswith(b'GIF87a') or content.startswith(b'GIF89a'):
                        mime_type = "image/gif"
                        filename = f"{id}.gif"
                    elif content.startswith(b'%PDF'):
                        mime_type = "application/pdf"
                        filename = f"{id}.pdf"
                    else:
                        # Check if it's text-based content (CSV, JSON, TXT, etc.)
                        try:
                            # Try to decode as UTF-8 text
                            text_content = content.decode('utf-8')

                            # Check if it looks like CSV
                            if ',' in text_content and ('\n' in text_content or '\r' in text_content):
                                # Simple heuristic: if it has commas and newlines, likely CSV
                                lines = text_content.split('\n')[:5]  # Check first 5 lines
                                if any(',' in line for line in lines):
                                    mime_type = "text/csv"
                                    filename = f"{id}.csv"
                                else:
                                    mime_type = "text/plain"
                                    filename = f"{id}.txt"
                            elif text_content.strip().startswith(('{', '[')):
                                # Looks like JSON
                                mime_type = "application/json"