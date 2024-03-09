import os
import pptx
from html2text import html2text
import json
import re

def get_filepaths(directory):
    file_paths = []
    for root, directories, files in os.walk(directory):
        if "Theory" in root:
            for filename in files:
                filepath = os.path.join(root, filename)
                file_paths.append(filepath)
    return file_paths


files = get_filepaths(r"C:\Users\Troy\Algonquin College")


def main():
    for file in files:
        extension = file.split(".")[-1]
        if extension == "pptx":
            #print(file)
            pptx_presentation = get_pptx_presentation(file)
            #print(pptx_presentation)

def get_pptx_presentation(file):
    presentation_dict = {"title":"","slides":[]}
    try:
        prs = pptx.Presentation(file)
        pptx_slides = get_pptx_slides(prs.slides)
        return pptx_slides
    except pptx.exc.PackageNotFoundError:
        print("File is not a pptx file")
        return None

def get_pptx_slides(presentation):
    slide_list = []
    for slide in presentation:
        slide_content = get_pptx_slide_content(slide)
        slide_list.append(slide_content)
    return slide_list

def get_pptx_slide_content(slide):
    slide_dict = {
        "title":"",
        "content":[]
    }
    try:
        title = slide.shapes.title.text
    except AttributeError:
        title = ""
        for shape in slide.shapes:
            if shape.has_text_frame:
                title = shape.text
                break
    # using regex, find any unicode characters in the title that are outside the ASCII range and convert the title to a list of strings
    regex = re.split(r'[^\x20-\x7E]', title)
    _slide_list = [x for x in regex if x != ""]
    if not _slide_list:
        pass
    else:
        if len(_slide_list) == 1:
            slide_dict["title"] = _slide_list[0]
        else:
            if len(_slide_list) == 2:
                if _slide_list[0][-1].islower() and _slide_list[1][0].isupper():
                    print(_slide_list)
if __name__ == "__main__":
    main()
