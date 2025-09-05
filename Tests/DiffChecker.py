import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import difflib

def select_file(entry_widget):
    """Open file dialog and set path in entry."""
    filepath = filedialog.askopenfilename(title="Select a file")
    if filepath:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, filepath)

def compare_files():
    file1 = entry1.get()
    file2 = entry2.get()

    if not file1 or not file2:
        messagebox.showerror("Error", "Please select both files.")
        return

    try:
        with open(file1, 'r') as f1, open(file2, 'r') as f2:
            lines1 = f1.readlines()
            lines2 = f2.readlines()
            
            differ = difflib.Differ()
            diff = list(differ.compare(lines1, lines2))
            
            changes = [line for line in diff if line.startswith('- ') or line.startswith('+ ')]
            
            # Classification
            if len(changes) < 50:
                classification = "Low Difference"
            elif len(changes) < 20:
                classification = "Medium Difference"
            else:
                classification = "High Difference"

            # Show result in text box
            text_area.delete(1.0, tk.END)
            text_area.insert(tk.END, "\n".join(diff))
            status_label.config(text=f"Classification: {classification}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to compare files.\n{e}")

# Tkinter UI
root = tk.Tk()
root.title("File Difference Checker")
root.geometry("800x600")

# File selectors
frame = tk.Frame(root)
frame.pack(pady=10)

entry1 = tk.Entry(frame, width=50)
entry1.grid(row=0, column=0, padx=5)
btn1 = tk.Button(frame, text="Browse File 1", command=lambda: select_file(entry1))
btn1.grid(row=0, column=1, padx=5)

entry2 = tk.Entry(frame, width=50)
entry2.grid(row=1, column=0, padx=5)
btn2 = tk.Button(frame, text="Browse File 2", command=lambda: select_file(entry2))
btn2.grid(row=1, column=1, padx=5)

compare_btn = tk.Button(root, text="Compare", command=compare_files, bg="lightblue")
compare_btn.pack(pady=10)

# Scrollable text area
text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=100, height=25)
text_area.pack(pady=10)

# Status label
status_label = tk.Label(root, text="Classification: N/A", font=("Arial", 12, "bold"))
status_label.pack(pady=5)

root.mainloop()
