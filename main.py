import os
import pptx
from html2text import html2text
import json


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
            pptx_file = PPTXfile(file)
            presentation = pptx_file.pptx_presentation
            print(presentation)

class PPTXfile:
    def __init__(self, file):
        self.file = file
        self.pptx_presentation = self.get_pptx_presentation(file)
        self.title = self.pptx_presentation["title"]
        self.slides = self.pptx_presentation["slides"]
    def get_pptx_presentation(self,file):
        presentation_dict ={"title": "", "slides": []}
        subject_items = file.split("\\")[5].split(" ",1)
        try:
            prs = pptx.Presentation(file)
            slides = self.get_pptx_slides(prs.slides)
            title_slide = slides[0]
            for slide in slides:
                if slide.get("title") == "":
                    # inherit the title from the previous slide
                    slide["title"] = title_slide["title"]
            if subject_items[-1] == title_slide["title"] or subject_items[0] == title_slide["title"]:
                if len(title_slide["content"]) > 0:
                    presentation_dict["title"] = title_slide["content"][0]
                else:
                    presentation_dict["title"] = title_slide["title"]
                presentation_dict["slides"] = slides[1:]
                return presentation_dict
            else:
                presentation_dict["title"] = title_slide["title"]
                presentation_dict["slides"] = slides[1:]
                return presentation_dict
        except pptx.exc.PackageNotFoundError:
            print(f"Error: {file} is not a valid pptx file")
            return None

    def get_pptx_slides(self,presentation):
        slides = []
        for slide in presentation:
            slide_content = self.get_pptx_slide_content(slide)
            slides.append(slide_content)
        return slides

    def get_pptx_slide_content(self,slide):
        slide_dict = {"title": "", "content": [],"images": []}
        try:
            title = slide.shapes.title.text
        except AttributeError:
            title = ""
            # if the slide does not have a title, get the title from the first shape that has text
            for shape in slide.shapes:
                if shape.has_text_frame:
                    title = shape.text
                    break
        invalid_chars = ["\"","/",":","*","?","<",">","|"]
        for char in invalid_chars:
            title = title.replace(char,"")
        if title.count("\x0b") > 0:
            _title_list = [html2text(x).strip() for x in list(filter(None, title.split("\x0b")))]
            if len(_title_list) ==2:
                slide_section = _title_list[0]
                slide_title = _title_list[-1]
                slide_dict["title"] = slide_title
                slide_dict["section"] = slide_section
            elif len(_title_list) == 1:
                slide_title = _title_list[0]
                slide_dict["title"] = slide_title
            elif len(_title_list) > 2:
                _title_list = " ".join(_title_list)
                slide_title = _title_list
                slide_dict["title"] = slide_title
        else:
            slide_dict["title"] = title
        for shape in slide.shapes:
            image_count = 0
            if shape.has_text_frame:
                if shape.text == title:
                    continue
                else:
                    for paragraph in shape.text_frame.paragraphs:
                        if len(html2text(paragraph.text).strip()) > 2:
                            slide_dict["content"].append(html2text(paragraph.text).strip())
                        elif paragraph.text.strip() == "":
                            continue
                        elif len(paragraph.text) <=2:
                            continue
            elif shape.shape_type == 19:
                table = shape.table
                table_content = ""
                # convert table to markdown, including headers
                for row_index, row in enumerate(table.rows):
                    if row_index == 0:
                        table_content += "|"
                        for cell in row.cells:
                            table_content += f"{cell.text}|"
                        table_content += "\n|"
                        for cell in row.cells:
                            table_content += "---|"
                        table_content += "\n"
                    else:
                        table_content += "|"
                        for cell in row.cells:
                            table_content += f"{cell.text}|"
                        table_content += "\n"
                slide_dict["content"].append(table_content)
            elif shape.shape_type == 13:
                image_count += 1
            elif shape.shape_type == 6:
                if shape.has_text_frame:
                    if shape.text == "":
                        continue
                    else:
                        slide_dict["content"].append(html2text(shape.text).strip())
                else:
                    continue
            elif shape.shape_type == 14:
                if shape.image:
                    image_count += 1
                else:
                    print(shape)
            elif shape.shape_type == 9:
                continue
            elif shape.shape_type is None:
                continue
            else:
                continue
        return slide_dict

if __name__ == "__main__":
    main()
