import os
import pptx
import re
from html2text import html2text


def get_filepaths(directory):
    file_paths = []
    for root, directories, filenames in os.walk(directory):
        if "Theory" in root:
            for filename in filenames:
                filepath = os.path.join(root, filename)
                file_paths.append(filepath)
    return file_paths


files = get_filepaths(r"D:\Algonquin College")


def main():
    extensions = []
    for file in files:
        extension = file.split(".")[-1]
        subject = file.split("\\")[3]
        if extension == "pptx":
            pptx_file = get_pptx_presentation(file)
            print(pptx_file)
        else:
            continue
        #else:
        #    pass

def get_pptx_slide_content(slide):
    slide_content = {"title": "", "content": []}
    try:
        _slide_title = slide.shapes.title.text
    except AttributeError:
        _slide_title = ""
        for shape in slide.shapes:
            if shape.has_text_frame:
                if shape.text.strip() == "":
                    continue
                else:
                    _slide_title = shape.text
                    break
            else:
                continue

    cisco_title_regex = re.match(r"^\S+(?: \S+){0,10}\x0b\S+(?: \S+){0,10}$", _slide_title)
    if cisco_title_regex:
        _title_list = _slide_title.split("\x0b")
        if _title_list[0][-1].islower() and _title_list[1][0].islower():
            _title_list = " ".join(_title_list)
        elif _title_list[0][-1] == ":" and _title_list[1][0].isupper():
            _title_list = _title_list[0] + " " + _title_list[1]
        else:
            slide_title = _title_list[-1]
            slide_section = _title_list[0]
            slide_content["title"] = slide_title
            slide_content["section"] = slide_section
    else:
        slide_content["title"] = html2text(_slide_title).strip()
    for shape in slide.shapes:
        if shape.has_text_frame:
            if shape.text.strip() == "":
                continue
            elif shape.text == _slide_title:
                continue
            else:
                for paragraph in shape.text_frame.paragraphs:
                    if paragraph.text.strip() == "":
                        continue
                    elif len(paragraph.text.strip()) <= 2 and paragraph.text.strip().isnumeric():
                        continue
                    else:
                        if paragraph.level > 0:
                            slide_content["content"].append(" "*paragraph.level + "- " + paragraph.text)
                        else:
                            slide_content["content"].append(paragraph.text)
        elif shape.has_table:
            # get the table as a list of lists
            table = []
            for row in shape.table.rows:
                table_row = []
                for cell in row.cells:
                    if cell.text.strip() == "":
                        continue
                    else:
                        table_row.append(" ".join(cell.text.split()))
                table.append(table_row)
            # append the table to the slide content as a Markdown table
            table_content = ""
            for i in range(len(table)):
                # if it is the first row, then add the header and a separator
                if i == 0:
                    table_content += "|"
                    for j in range(len(table[i])):
                        table_content += table[i][j] + "|"
                    table_content += "\n"
                    table_content += "|"
                    for j in range(len(table[i])):
                        table_content += "---|"
                    table_content += "\n"
                else:
                    table_content += "|"
                    for j in range(len(table[i])):
                        table_content += table[i][j] + "|"
                    table_content += "\n"
            slide_content["content"].append(table_content)
    return slide_content

def get_pptx_slides(presentation):
    pptx_slides = []
    for slide in presentation.slides:
        pptx_slide_content = get_pptx_slide_content(slide)
        if pptx_slide_content["title"] == "" or len(pptx_slide_content["content"]) == 0:
            continue
        else:
            pptx_slides.append(pptx_slide_content)
    return pptx_slides


def get_pptx_presentation(file):
    presentation_dict = {"title": "", "slides": []}
    subject = file.split("\\")[3]
    presentation_dict["subject"] = subject
    subject_items = subject.split(" ", 1)
    try:
        prs = pptx.Presentation(file)
        pptx_slides = get_pptx_slides(prs)
        title_slide = pptx_slides[0]
        if title_slide["title"] == subject_items[-1] or title_slide["title"] == subject_items[0]:
            title_slide["title"] = title_slide["content"][0]
            title_slide["content"].pop(0)
        else:
            pass
        presentation_dict["title"] = title_slide["title"]
        presentation_dict["slides"] = pptx_slides[1:]
    except pptx.exc.PackageNotFoundError:
        return None
    return presentation_dict

class ObsidianVault:
    def __init__(self, path):
        self.path = path

    def create_folder(self, folder_path):
        if not os.path.exists(os.path.join(self.path, folder_path)):
            os.mkdir(os.path.join(self.path, folder_path))
        else:
            pass

    def create_file(self, file_path, content):
        try:
            with open(os.path.join(self.path, file_path), "w", encoding="utf-8") as file:
                file.write(content)
        except OSError:
            pass


if __name__ == "__main__":
    main()
