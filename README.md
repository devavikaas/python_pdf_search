# python_pdf_search
This Python script scans PDF files in a selected folder (and subfolders) for user-defined keywords, extracts surrounding context lines, and saves results to Excel. It includes a simple GUI for input, uses PyMuPDF for text extraction, supports multiprocessing for speed, and highlights matches in the terminal. Ideal for analyzing large PDF sets.

# PDF Search Tool with GUI, Highlighting, and Excel Export
# Line-by-line Explanation in Notepad Style

* import os
  - Used to navigate file paths and walk through directories

* import re
  - For compiling and using regular expressions to match search terms

* import fitz  # PyMuPDF
  - Main library to read and process PDF files

* import pandas as pd
  - Used for creating and saving structured data (Excel output)

* import tkinter as tk
  - GUI framework to prompt user inputs

* from tkinter import filedialog, simpledialog
  - Filedialog: to select folders
  - Simpledialog: to prompt for search terms and context lines

* from datetime import datetime
  - Used to generate timestamped output filenames

* from tqdm import tqdm
  - For showing progress bars while processing PDFs

* from multiprocessing import Pool, cpu_count
  - Enables parallel processing using multiple CPU cores

# === GUI Setup ===

* def select_folder():
  - Opens a folder picker dialog and returns the selected folder path

* def get_search_terms():
  - Opens a dialog to enter comma-separated search terms
  - Returns a list of trimmed search terms

* def get_line_context():
  - Asks user how many lines before/after a match to display
  - Uses try/except to default to 1 if user input fails
  - Returns two integers (before, after)

# === Highlight Matched Term in Terminal Output ===

* def highlight_term(text, term):
  - Uses regex to find term occurrences in the line
  - Wraps matched term in ANSI color codes for red bold
  - Makes terminal output easier to read

# === Worker Function to Search a Single File ===

* def search_pdf(args):
  - Takes a tuple: (file path, search terms, lines before, lines after)
  - Returns a list of result dictionaries

  - Extracts text from each page of the PDF
  - Splits text into lines
  - For each search term:
    - Finds line indexes that contain the term
    - Gathers context lines before/after the match
    - Creates two versions:
      - Raw context (for Excel)
      - Highlighted context (for terminal)

  - Appends the result dictionary:
    - File Name, Page Number, Matched Term, Count, Context, Console Context

  - Handles exceptions gracefully and prints errors

# === Save Results to Excel ===

* def save_results_to_excel(results, output_file):
  - Converts results to a DataFrame
  - Drops "Console Context" (used only for printing)
  - Sorts by file name, page number, and term
  - Saves the cleaned results as an Excel (.xlsx) file

# === Main Program ===

* def main():
  - Main controller function for the script

  - folder_path = select_folder()
    - Prompts the user to select a folder with PDFs

  - search_terms = get_search_terms()
    - Prompts for comma-separated terms

  - lines_before, lines_after = get_line_context()
    - Asks how much context to include before/after match lines

  - Recursively walks through the folder to find all .pdf files

  - Prepares a list of args for multiprocessing

  - Uses multiprocessing.Pool with all available CPU cores
    - Searches all PDFs in parallel using search_pdf()

  - Flattens nested results list

  - If results found:
    - Prints each match with highlighted context to console
    - Builds a clean filename based on search terms and timestamp
    - Saves to Excel

  - Else: print "No matches found."

# === Script Entry Point ===

* if __name__ == "__main__":
  - Runs main() only if this file is executed directly

