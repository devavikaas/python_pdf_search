import os
import re
import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog
from datetime import datetime
from tqdm import tqdm
from multiprocessing import Pool, cpu_count

# === GUI Setup ===
def select_folder():
    root = tk.Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory(title="Select Folder Containing PDFs")
    root.destroy()
    return folder_selected

def get_search_terms():
    root = tk.Tk()
    root.withdraw()
    input_text = simpledialog.askstring("Search Terms", "Enter search terms or phrases separated by commas:")
    root.destroy()
    return [term.strip() for term in input_text.split(",")] if input_text else []

def get_line_context():
    root = tk.Tk()
    root.withdraw()
    try:
        lines_before = simpledialog.askinteger("Context Lines", "How many lines BEFORE a match should be shown?", minvalue=0, initialvalue=1)
        lines_after = simpledialog.askinteger("Context Lines", "How many lines AFTER a match should be shown?", minvalue=0, initialvalue=1)
    except Exception:
        lines_before, lines_after = 1, 1
    root.destroy()
    return lines_before or 0, lines_after or 0

# === Highlight Matched Term in Terminal Output ===
def highlight_term(text, term):
    pattern = re.compile(rf"(\b{re.escape(term)}\b)", re.IGNORECASE)
    return pattern.sub(r"\033[1;31m\1\033[0m", text)

# === Worker Function to Search a Single File ===
def search_pdf(args):
    pdf_path, search_terms, lines_before, lines_after = args
    results = []
    file_name = os.path.abspath(pdf_path)
    term_patterns = {term: re.compile(rf"\b{re.escape(term)}\b", re.IGNORECASE) for term in search_terms}

    try:
        doc = fitz.open(pdf_path)
        for i in range(len(doc)):
            page = doc.load_page(i)
            text = page.get_text("text")
            if not text:
                continue

            lines = text.splitlines()
            for term, pattern in term_patterns.items():
                match_indexes = [idx for idx, line in enumerate(lines) if pattern.search(line)]
                if not match_indexes:
                    continue

                context_blocks_raw = []
                context_blocks_highlighted = []

                for idx in match_indexes:
                    start = max(idx - lines_before, 0)
                    end = min(idx + lines_after + 1, len(lines))
                    block_raw = lines[start:end]
                    block_highlighted = [highlight_term(line, term) for line in block_raw]

                    context_blocks_raw.append("• " + "\n".join(line.strip() for line in block_raw))
                    context_blocks_highlighted.append("• " + "\n".join(line.strip() for line in block_highlighted))

                results.append({
                    "File Name": file_name,
                    "Page Number": i + 1,
                    "Matched Term": term,
                    "Count": len(match_indexes),
                    "Context": "\n\n".join(context_blocks_raw),
                    "Console Context": "\n\n".join(context_blocks_highlighted)
                })
        doc.close()
    except Exception as e:
        print(f"❌ Error reading {file_name}: {e}")
    return results

# === Save Results to Excel ===
def save_results_to_excel(results, output_file):
    df = pd.DataFrame(results)
    df.drop(columns=["Console Context"], inplace=True, errors="ignore")
    df.sort_values(by=["File Name", "Page Number", "Matched Term"], inplace=True)
    df.to_excel(output_file, index=False)

# === Main Program ===
def main():
    folder_path = select_folder()
    if not folder_path:
        print("No folder selected. Exiting.")
        return

    search_terms = get_search_terms()
    if not search_terms:
        print("No search terms entered. Exiting.")
        return

    lines_before, lines_after = get_line_context()

    print(f"\n🔍 Searching for: {', '.join(search_terms)}")
    print(f"📂 Folder: {folder_path} (including subfolders)")
    print(f"📌 Context: {lines_before} line(s) before and {lines_after} line(s) after matches\n")

    pdf_files = [
        os.path.join(root, file)
        for root, _, files in os.walk(folder_path)
        for file in files
        if file.lower().endswith(".pdf")
    ]
    total_files = len(pdf_files)

    if total_files == 0:
        print("❌ No PDF files found in the selected folder or its subfolders.")
        return

    print(f"⚙️ Found {total_files} PDF(s) — Using {cpu_count()} CPU cores...\n")

    args = [(file, search_terms, lines_before, lines_after) for file in pdf_files]

    with Pool(cpu_count()) as pool:
        all_results = list(tqdm(pool.imap(search_pdf, args),
                                total=total_files, desc="Overall Progress", unit="file"))

    flat_results = [item for sublist in all_results for item in sublist]

    if flat_results:
        for result in flat_results:
            print(f"\n✅ '{result['Matched Term']}' in {result['File Name']} (Page {result['Page Number']})")
            print(f"   Context:\n{result['Console Context']}\n")

        clean_terms = "_".join(re.sub(r"[^\w\-]+", "", term) for term in search_terms)
        output_file = os.path.join(folder_path, f"{clean_terms}_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

        save_results_to_excel(flat_results, output_file)
        print(f"\n📁 Results saved to: {output_file}")
    else:
        print("❌ No matches found.")

# === Run Script ===
if __name__ == "__main__":
    main()
