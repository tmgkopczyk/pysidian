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
        if any(slide_content.values()):
            if slide_content.get("title") == "":
                try:
                    slide_content["title"] = slide_content["content"][0]
                    slide_content["content"].pop(0)
                except IndexError:
                    pass
            print(slide_content)
            slide_list.append(slide_content)
        else:
            pass
    return slide_list

def get_pptx_slide_content(slide):
    slide_dict = {
        "title":"",
        "content":[],
        "images":[]
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
    regex = re.split(r"[\x0b+\xa0]",title)
    _slide_list = [x for x in regex if x != ""]
    if not _slide_list:
        pass
    else:
        if len(_slide_list) == 2:
            if (_slide_list[0][-1].islower() or _slide_list[0][-1].isupper() or (_slide_list[0][-1].isdigit() and _slide_list[1][0].isdigit() is False)) and (_slide_list[1][0].isupper() or _slide_list[1][0].isdigit()):
                slide_dict["section"] = _slide_list[0]
                slide_dict["title"] = _slide_list[-1]
            elif _slide_list[0][-1] == " " or _slide_list[0][-1] == ":" or _slide_list[0][-1] == ",":
                _slide_list = " ".join([x.strip() for x in _slide_list])
                slide_dict["title"] = _slide_list
            elif _slide_list[0][-1].islower() and _slide_list[1][0].islower():
                _slide_list = " ".join([x.strip() for x in _slide_list if x.strip() != ""])
                slide_dict["title"] = _slide_list
            elif _slide_list[-1][0] == " " and _slide_list[0][-1] != " ":
                _slide_list = " ".join([x.strip() for x in _slide_list if x.strip() != ""])
                slide_dict["title"] = _slide_list
            else:
                _slide_list = " ".join([x.strip() for x in _slide_list])
                slide_dict["title"] = _slide_list
        elif len(_slide_list) == 1:
            slide_dict["title"] = _slide_list[0]
        elif len(_slide_list) > 2:
            if _slide_list[-1][0].islower():
                slide_dict["title"] = " ".join([x.strip() for x in _slide_list if x.strip() != ""])
            elif _slide_list[0][-1] == ":":
                slide_dict["title"] = " ".join([x.strip() for x in _slide_list if x.strip() != ""])
            # if the first character of each item in the list is an uppercase letter, then join the list into a string
            elif _slide_list[0][0].isupper() and _slide_list[1][0].isupper():
                slide_dict["title"] = " ".join([x.strip() for x in _slide_list if x.strip() != ""])
            elif _slide_list[-2] == "–" or _slide_list[-2].islower():
                slide_dict["title"] = " ".join([x.strip() for x in _slide_list if x.strip() != ""])
            else:
                slide_dict["title"] = " ".join([x.strip() for x in _slide_list if x.strip() != ""])
    image_count = 0
    for shape in slide.shapes:
        if shape.has_text_frame:
            if shape.text == title:
                pass
            else:
                for paragraph in shape.text_frame.paragraphs:
                    paragraph_text = ""
                    if len(paragraph.text.strip()) <=2:
                        pass
                    elif paragraph.text.strip() == "":
                        pass
                    else:
                        for run in paragraph.runs:
                            if run.font.name is None:
                                # get the font from the slide master
                                for shape_l in slide.slide_layout.slide_master.slide_layouts.get_by_name(slide.slide_layout.name).shapes:
                                    if shape_l.shape_type == shape.shape_type:
                                        for paragraph_l in shape_l.text_frame.paragraphs:
                                            for run_l in paragraph_l.runs:
                                                print(run_l.font.bold)


        elif shape.shape_type == 13:
            image_count += 1
            slide_dict["images"].append(f"{slide_dict['title'].lower().replace(' ','_')}_{image_count}")
    return slide_dict

if __name__ == "__main__":
    main()
