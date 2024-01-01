import os
import pptx
from html2text import html2text
import json

def get_filepaths(directory):
    file_paths = []
    for root, directories, files in os.walk(directory):
        if "Slides" in root:
            for filename in files:
                filepath = os.path.join(root, filename)
                file_paths.append(filepath)
    return file_paths

files = get_filepaths(r"C:\Users\Troy\Algonquin\Fall 2023")

def get_slides(presentation_slides,subject):
    if subject == "Networking Fundamentals":
        slides = []
        section_dict = {
            "section": "",
            "slides": []
        }
        for slide_index, slide in enumerate(presentation_slides):
            slide_content = get_slide_content(slide,subject)
            # if section key in slide_content, then it is a new section
            if "section" in slide_content:
                # if the section is not empty, then add it to the slides
                if section_dict["section"] != "":
                    slides.append(section_dict)
                # create a new section
                section_dict = slide_content
            else:
                section_dict["slides"].append(slide_content)
            # if we are on the last slide, then add the last section
        slides.append(section_dict)
        return slides


def get_slide_content(slide,subject):
    if subject == "Networking Fundamentals":
        if slide.slide_layout.name == "3_Segue":
            # begin new section
            section_dict = {
                "section":"",
                "slides":[]
            }

            # get the section title from the title of the slide
            try:
                section_title = slide.shapes.title.text
            except AttributeError:
                section_title = ""
            section_dict["section"] = section_title
            return section_dict
        else:
            slide_dict = {
                "title":"",
                "content":[]
            }
            # get the title of the slide
            try:
                title = slide.shapes.title.text
            except AttributeError:
                title = ""
            slide_dict["title"] = title
            for shape in slide.shapes:
                try:
                    if shape.has_text_frame:
                        if shape.text == slide_dict["title"]:
                            continue
                        else:
                            try:
                                shape.fill.solid()
                                if shape.fill.fore_color.rgb[0] == 0 and shape.fill.fore_color.rgb[1] == 0 and shape.fill.fore_color.rgb[2] == 0 or shape.fill.fore_color.rgb[0] == 0 and shape.fill.fore_color.rgb[1] == 176 and shape.fill.fore_color.rgb[2] == 80 or \
                                        shape.fill.fore_color.rgb[0] == 8 and shape.fill.fore_color.rgb[1] == 8 and shape.fill.fore_color.rgb[2] == 8:
                                    slide_dict["content"].append("```")
                                    slide_dict["content"].append(shape.text)
                                    slide_dict["content"].append("```")
                                else:
                                    continue
                            except AttributeError:
                                # append as regular text
                                for paragraph in shape.text_frame.paragraphs:
                                    try:
                                        if "(" in paragraph.text[0] and ")" in paragraph.text[-1]:
                                            continue
                                        elif paragraph.text.strip() == "" or len(paragraph.text.strip()) <= 3:
                                            continue
                                        else:
                                            # append paragraph text to slide content based on level
                                            if paragraph.level == 0:
                                                slide_dict["content"].append(paragraph.text)
                                            elif paragraph.level > 0:
                                                slide_dict["content"].append(" " * paragraph.level + "- " + paragraph.text)
                                    except IndexError:
                                        continue
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
                        # append the table to the slide content as a markdown table
                        table_content = ""
                        table_content += "\n"
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
                        slide_dict["content"].append(table_content)
                except AttributeError:
                    continue
            return slide_dict
def handle_networking_fundamentals(file_path):
    presentation = {}
    prs = pptx.Presentation(file_path)
    # get the title of the presentation from the title of the first slide
    try:
        presentation_title = prs.slides[0].shapes.title.text
    except AttributeError:
        presentation_title = ""
    presentation["title"] = presentation_title
    slides = get_slides(prs.slides,"Networking Fundamentals")
    presentation["slides"] = slides

    return presentation

def main():
    presentations = []
    for file in files:
        if "Networking Fundamentals" in file:
            networking_presentation = handle_networking_fundamentals(file)
            create_presentation_folder("Networking Fundamentals",networking_presentation)

class ObsidianVault:
    def __init__(self,path):
        self.path = path

    def create_folder(self,folder_path):
        if not os.path.exists(os.path.join(self.path,folder_path)):
            os.mkdir(os.path.join(self.path,folder_path))
        else:
            pass
vault = ObsidianVault(r"C:\Users\Troy\Obsidian\College")

def create_presentation_folder(subject, presentation):
    invalid_chars = ["\\","/",":","*","?","\"","<",">","|"]
    # create a folder for the presentation inside the subject folder
    presentation_folder_name = presentation["title"]
    presentation_folder_name = presentation_folder_name.split(": ")[1]
    try:
        os.mkdir(os.path.join(vault.path,subject))
    except FileExistsError:
        pass
    vault.create_folder(os.path.join(subject,presentation_folder_name))
    # create a folder for each section inside the presentation folder
    for section in presentation["slides"]:
        section_folder_name = section["section"].split(" ",1)[1]
        section_path = os.path.join(subject,presentation_folder_name,section_folder_name)
        vault.create_folder(section_path)
        # create a markdown file for each slide inside the section folder
        for slide in section["slides"]:
            if "\x0b" in slide["title"]:
                slide["title"] = slide["title"].split("\x0b")[1]
            slide_file_name = slide["title"]
            for char in invalid_chars:
                slide_file_name = slide_file_name.replace(char,"")
            slide_file_name = slide_file_name + ".md"
            try:
                with open(str(os.path.join(str(vault.path),str(section_path),str(slide_file_name))),"w",encoding="utf-8") as slide_file:
                    for content in slide["content"]:
                        slide_file.write(content)
                        slide_file.write("\n")

            except OSError:
                continue
if __name__ == "__main__":
    main()

