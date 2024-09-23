#Script to find pdf files with keywords in them
import os
import sys
import PyPDF2

def find_keyword_in_pdfs(folder_path, keyword):
    matching_files = []

    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf'):
            try:
                file_path = os.path.join(folder_path, filename)
                with open(file_path, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    num_pages = len(pdf_reader.pages)

                    for page_num in range(num_pages):
                        page = pdf_reader.pages[page_num]
                        text = page.extract_text()
                        if keyword in text:
                            matching_files.append(filename)
                            break  # Stop searching this file as soon as the keyword is found
            except Exception as e:
                print(f"Error reading {filename}: {e}")

    return matching_files

#TODO: Add an API call to Sabiá-3 to get a summary of the files that contain the keyword
def get_sabia_summary(folder_path, keyword):
    pass

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python pdf_keyword_search.py [folder_path] [keyword]")
        sys.exit(1)

    folder_path = sys.argv[1]
    keyword = sys.argv[2]
    matching_files = find_keyword_in_pdfs(folder_path, keyword)

    print("Files containing the keyword:")
    for file in matching_files:
        print(file)
