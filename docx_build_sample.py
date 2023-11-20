import os
from docx import Document
from dotenv import load_dotenv
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Mm


def main():
    user = os.getenv("USERNAME")
    file = Document()

    # ファイルの余白を設定
    section = file.sections[0]
    section.top_margin = Mm(30)
    section.bottom_margin = Mm(20)
    section.left_margin = Mm(20)
    section.right_margin = Mm(16)

    paragraph_format = file.styles["Normal"].paragraph_format
    # デフォルトの段落スペースを0ポイントに設定
    paragraph_format.space_before = Pt(0)
    paragraph_format.space_after = Pt(0)

    p = file.add_paragraph("報告日：2023/xx/xx")
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.right_indent = Pt(30)

    file.add_paragraph()
    p = file.add_paragraph("学籍番号：xxxxxx")
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.right_indent = Pt(75)
    p = file.add_paragraph("氏名：xxxxxx")
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.right_indent = Pt(75)

    file.save("/mnt/c/Users/" + user + "/Downloads/sample.docx")


if __name__ == "__main__":
    load_dotenv()
    main()
