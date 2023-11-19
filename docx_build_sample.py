import os
from docx import Document
from dotenv import load_dotenv


def main():
    user = os.getenv("USERNAME")
    file = Document()

    file.add_heading("Sample Document", 0)
    file.add_paragraph("This is a sample paragraph.")
    file.add_paragraph("This is another sample paragraph.")
    file.add_page_break()
    file.add_paragraph("This is a paragraph on the second page.")
    file.save("sample.docx")
    file.save("/mnt/c/Users/" + user + "/Downloads/sample.docx")


if __name__ == "__main__":
    load_dotenv()
    main()
