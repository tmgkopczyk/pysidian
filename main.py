import os
import pptx
import re
import pypdf
from pypdf.errors import PdfReadError


def get_filepaths(directory):
    file_paths = []
    for root, directories, filenames in os.walk(directory):
        if "Theory" in root:
            for filename in filenames:
                filepath = os.path.join(root, filename)
                file_paths.append(filepath)
    return file_paths


files = get_filepaths(r"C:\Users\Troy\Algonquin College")


def main():
    extensions = []
    for file in files:
        extension = file.split(".")[-1]
        if extension == "pptx":
            pptx_file = get_pptx_presentation(file)
            #print(pptx_file)
            print(file)
        else:
            continue

def get_pptx_slide_content(slide):
    slide_content = {"title": "", "content": []}
    try:
        slide_title = slide.shapes.title.text
    except AttributeError:
        slide_title = ""
        for shape in slide.shapes:
            if shape.has_text_frame:
                slide_title = shape.text
                break
            else:
                continue
    slide_title = slide_title.strip()
    regex = re.match(r"^\S+(?: \S+){0,10}\x0b\S+(?: \S+){0,10}$", slide_title)
    if regex:
        _title_list = [x for x in slide_title.split("\x0b") if x.strip() != ""]
        if _title_list[0][-1].islower() and _title_list[1][0].islower():
            slide_title = " ".join([x.strip() for x in _title_list])
        elif _title_list[0][-1] == ":" and _title_list[1][0].isupper():
            slide_title = " ".join([x.strip() for x in _title_list])
        else:
            slide_content["title"] = _title_list[-1]
            slide_content["section"] = _title_list[0]
    else:
        _slide_title = re.findall("\x0b", slide_title)
        if len(_slide_title) == 1:
            slide_title = " ".join([x.strip() for x in slide_title.split("\x0b") if x.strip() != ""])
        else:
            _slide_title = [x.strip() for x in slide_title.splitlines() if x.strip() != ""]
            if len(_slide_title) > 1:
                if [x[0].strip().islower() for x in _slide_title[1:]] == [True] * (len(_slide_title) - 1):
                    _slide_title = " ".join([x.strip() for x in _slide_title])
                else:
                    print(_slide_title)
        slide_content["title"] = slide_title
    for shape in slide.shapes:
        if shape.has_text_frame:
            if shape.text == slide_title:
                continue
            else:
                for paragraph in shape.text_frame.paragraphs:
                    if len(paragraph.text.strip()) <= 2 and paragraph.text.isnumeric():
                        continue
                    elif paragraph.text.strip() == "":
                        continue
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
    if slide_content["title"] == "":
        if len(slide_content["content"]) > 0:
            slide_content["title"] = slide_content["content"][0]
            slide_content["content"].pop(0)
        else:
            slide_content["title"] = "No title"

    return slide_content


def get_pptx_slides(presentation):
    pptx_slides = []
    for slide in presentation.slides:
        pptx_slide_content = get_pptx_slide_content(slide)
        if pptx_slide_content["title"] == "No title" and len(pptx_slide_content["content"]) == 0:
            continue
        else:
            pptx_slides.append(pptx_slide_content)
    return pptx_slides


def get_pptx_presentation(file):
    presentation_dict = {"title": "", "slides": []}
    subject = file.split("\\")[5]
    subject_items = subject.split(" ", 1)
    try:
        prs = pptx.Presentation(file)
        pptx_slides = get_pptx_slides(prs)
        title_slide = pptx_slides[0]
        if title_slide["title"] == subject_items[0] or title_slide["title"] == subject_items[-1]:
            title_slide["title"] = title_slide["content"][0]
            title_slide["content"].pop(0)
        else:
            pass
        presentation_dict["title"] = title_slide["title"]
        presentation_dict["slides"] = pptx_slides[1:]
    except pptx.exc.PackageNotFoundError:
        return None
    return presentation_dict


if __name__ == "__main__":
    main()
