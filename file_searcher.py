import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from docx import Document
import json
from PyPDF2 import PdfReader
from pptx import Presentation


def check_file_contents(filename, search_string, case_sensitive=True, exact_match=False):
    file_extension = os.path.splitext(filename)[1]
    script_path = os.path.abspath(__file__)
    original_search_string = search_string

    if not case_sensitive:
        search_string = search_string.lower()

    if file_extension == '.docx':
        doc = Document(filename)
        for paragraph in doc.paragraphs:
            if not case_sensitive:
                paragraph_text = paragraph.text.lower()
            else:
                paragraph_text = paragraph.text

            if exact_match:
                if search_string == paragraph_text:
                    results_text.insert(tk.END,
                                        f"Found '{original_search_string}' in {filename} - Line: {paragraph.text}\n")
            else:
                if search_string in paragraph_text:
                    results_text.insert(tk.END,
                                        f"Found '{original_search_string}' in {filename} - Line: {paragraph.text}\n")

    elif file_extension == '.py':
        with open(filename, 'r') as file:
            for line_number, line in enumerate(file, start=1):
                if not case_sensitive:
                    line = line.lower()

                if exact_match:
                    if search_string == line.strip():
                        results_text.insert(tk.END,
                                            f"Found '{original_search_string}' in {filename} - Line {line_number}: {line.strip()}\n")
                else:
                    if search_string in line:
                        results_text.insert(tk.END,
                                            f"Found '{original_search_string}' in {filename} - Line {line_number}: {line.strip()}\n")


    elif file_extension == '.txt':
        with open(filename, 'r') as file:
            for line_number, line in enumerate(file, start=1):
                if not case_sensitive:
                    line = line.lower()

                if exact_match:
                    if search_string == line.strip():
                        results_text.insert(tk.END,
                                            f"Found '{original_search_string}' in {filename} - Line {line_number}: {line.strip()}\n")
                else:
                    if search_string in line:
                        results_text.insert(tk.END,
                                            f"Found '{original_search_string}' in {filename} - Line {line_number}: {line.strip()}\n")

    elif file_extension == '.json':
        with open(filename, 'r') as file:
            data = json.load(file)
            if not case_sensitive:
                data_str = str(data).lower()
            else:
                data_str = str(data)

            if exact_match:
                if search_string == data_str:
                    results_text.insert(tk.END, f"Found '{original_search_string}' in {filename}\n")
            else:
                if search_string in data_str:
                    results_text.insert(tk.END, f"Found '{original_search_string}' in {filename}\n")

    elif file_extension == '.pdf':
        with open(filename, 'rb') as file:
            pdf_reader = PdfReader(file)
            for page_number, page in enumerate(pdf_reader.pages, start=1):
                page_content = page.extract_text()
                if not case_sensitive:
                    page_content = page_content.lower()

                if exact_match:
                    if search_string == page_content.strip():
                        results_text.insert(tk.END,
                                            f"Found '{original_search_string}' in {filename} - Page {page_number}\n")
                else:
                    if search_string in page_content:
                        results_text.insert(tk.END,
                                            f"Found '{original_search_string}' in {filename} - Page {page_number}\n")

    elif file_extension == '.pptx':
        prs = Presentation(filename)
        for slide_number, slide in enumerate(prs.slides, start=1):
            slide_content = []
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    slide_content.append(shape.text)
            slide_text = '\n'.join(slide_content)

            if not case_sensitive:
                slide_text = slide_text.lower()

            if exact_match:
                if search_string == slide_text.strip():
                    results_text.insert(tk.END,
                                        f"Found '{original_search_string}' in {filename} - Slide {slide_number}\n")
            else:
                if search_string in slide_text:
                    results_text.insert(tk.END,
                                        f"Found '{original_search_string}' in {filename} - Slide {slide_number}\n")

    else:
        results_text.insert(tk.END, f"Unsupported file format: {file_extension}\n")


def browse_directory():
    selected_directory = filedialog.askdirectory()
    directory_var.set(selected_directory)


def search_files():
    search_string = search_entry.get().strip()
    if not search_string:
        messagebox.showwarning("Empty Search String", "Please enter a search string.")
        return

    case_sensitive = case_sensitive_var.get()
    exact_match = exact_match_var.get()

    search_directory = directory_var.get()
    if not search_directory:
        messagebox.showwarning("No Directory Selected", "Please select a directory to search.")
        return

    script_file = os.path.normpath(os.path.abspath(__file__))

    results_text.delete(1.0, tk.END)

    for filename in os.listdir(search_directory):
        file_path = os.path.normpath(os.path.join(search_directory, filename))

        if os.path.isfile(file_path) and not filename.startswith('~$') and file_path != script_file:
            check_file_contents(file_path, search_string, case_sensitive, exact_match)


# Create the main window
window = tk.Tk()
window.title("File Searcher")
window.geometry("500x400")

# Create and place widgets
search_label = tk.Label(window, text="Search String:")
search_label.pack()

search_entry = tk.Entry(window, width=50)
search_entry.pack()

case_sensitive_var = tk.BooleanVar()
case_sensitive_checkbox = tk.Checkbutton(window, text="Case Sensitive", variable=case_sensitive_var)
case_sensitive_checkbox.pack()

exact_match_var = tk.BooleanVar()
exact_match_checkbox = tk.Checkbutton(window, text="Exact Match", variable=exact_match_var)
exact_match_checkbox.pack()

directory_var = tk.StringVar()
directory_label = tk.Label(window, text="Search Directory:")
directory_label.pack()

directory_frame = tk.Frame(window)
directory_entry = tk.Entry(directory_frame, textvariable=directory_var, width=40)
directory_entry.pack(side=tk.LEFT)

browse_button = tk.Button(directory_frame, text="Browse", command=browse_directory)
browse_button.pack(side=tk.LEFT)

directory_frame.pack()

search_button = tk.Button(window, text="Search", command=search_files)
search_button.pack()

results_label = tk.Label(window, text="Results:")
results_label.pack()

results_text = tk.Text(window, width=60, height=15)
results_text.pack()

window.mainloop()
