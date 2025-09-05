import os
import shutil

# Path to your dataset directory
base_dir = r"D:\EXSQ\TextTopo\Data\TXT"
sorted_dir = os.path.join(base_dir, "Sorted")

# Define language mappings (lowercased for case-insensitivity)
language_map = {
    "sp": "Spanish",
    "es": "Spanish",
    "ar": "Arabic",
    "ab": "Arabic",
    "en": "English",
    "en sp": "English_Spanish",
    "en es": "English_Spanish"
}

def get_language(filename: str) -> str:
    name, _ = os.path.splitext(filename)
    parts = name.split()

    # Check last two tokens (for "EN SP", case-insensitive)
    last_two = " ".join(parts[-2:]).lower()
    if last_two in language_map:
        return language_map[last_two]

    # Check last token (case-insensitive)
    last = parts[-1].lower()
    if last in language_map:
        return language_map[last]

    # Default: English (if no suffix)
    return "English"

# Create a "Sorted" folder and copy files
os.makedirs(sorted_dir, exist_ok=True)

for file in os.listdir(base_dir):
    if file.endswith(".txt"):
        lang = get_language(file)
        target_dir = os.path.join(sorted_dir, lang)
        os.makedirs(target_dir, exist_ok=True)

        shutil.copy2(os.path.join(base_dir, file),
                     os.path.join(target_dir, file))

print("✅ Files have been copied into 'Sorted' subfolders by language!")
