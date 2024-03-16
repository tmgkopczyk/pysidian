import os
import pptx
import re


def get_filepaths(directory):
    file_paths = []
    for root, directories, filenames in os.walk(directory):
        if "Theory" in root:
            for filename in filenames:
                filepath = os.path.join(root, filename)
                file_paths.append(filepath)
    return file_paths


files = get_filepaths(r"C:\Users\Troy\Algonquin College")

def get_pptx_slides(pptx_presentation):
    slides = []
    for slide in pptx_presentation.slides:
        slide_content = get_pptx_slide_content(slide)
        slides.append(slide_content)
    return slides

def get_pptx_slide_content(slide):
    slide_dict = {
        "title": "",
        "content": [],
        "pictures": []
    }
    # get the title of the slide
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
    _slide_list = [x for x in list(filter(None, regex)) if x.strip() != ""]
    if not _slide_list:
        slide_dict["title"] = title
    else:
        if len(_slide_list) == 1:
            slide_dict["title"] = _slide_list[0]
        elif len(_slide_list) == 2:
            if (_slide_list[0][-1].islower() and _slide_list[-1][0].isupper()) or (_slide_list[0][-1].isupper() and _slide_list[-1][0].isupper()) or (_slide_list[0][-1].isnumeric() and _slide_list[-1][0].isupper()) or (_slide_list[0][-1].islower() and _slide_list[-1][0].isnumeric()):
                slide_dict["title"] = _slide_list[-1]
                slide_dict["section"] = _slide_list[0]
            elif (_slide_list[0][-1] == " " and _slide_list[-1][0].isupper()) or (_slide_list[0][-1] == " " and _slide_list[-1][0].islower()) :
                _slide_list = " ".join([x.strip() for x in _slide_list])
                slide_dict["title"] = _slide_list
            elif _slide_list[0][-1].islower() and _slide_list[-1][0].islower():
                _slide_list = " ".join([x.strip() for x in _slide_list])
                slide_dict["title"] = _slide_list
            elif _slide_list[0][-1] == "," or _slide_list[0][-1] == ":":
                _slide_list = " ".join([x.strip() for x in _slide_list])
                slide_dict["title"] = _slide_list
            else:
                _slide_list = " ".join([x.strip() for x in _slide_list])
                slide_dict["title"] = _slide_list
        elif len(_slide_list) > 2:
            if _slide_list[0][-1].islower() or _slide_list[0][-1].islower() and _slide_list[-1][0].islower():
                _slide_list = " ".join([x.strip() for x in _slide_list])
                slide_dict["title"] = _slide_list
            elif _slide_list[0][-1] == "," or _slide_list[0][-1] == ":":
                _slide_list = " ".join([x.strip() for x in _slide_list])
                slide_dict["title"] = _slide_list
            else:
                slide_dict["title"] = _slide_list[-1]
    #slide_dict["title"] = title
    for shape in slide.shapes:
        try:
            if shape.has_text_frame:
                if shape.text == slide_dict["title"]:
                    continue
                else:
                    try:
                        shape.fill.solid()
                        if shape.fill.fore_color.rgb[0] == 0 and shape.fill.fore_color.rgb[1] == 0 and \
                                shape.fill.fore_color.rgb[2] == 0 or shape.fill.fore_color.rgb[0] == 0 and \
                                shape.fill.fore_color.rgb[1] == 176 and shape.fill.fore_color.rgb[2] == 80 or \
                                shape.fill.fore_color.rgb[0] == 8 and shape.fill.fore_color.rgb[1] == 8 and \
                                shape.fill.fore_color.rgb[2] == 8:
                            if shape.text.strip() == "":
                                continue
                            else:
                                # append as code block
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
                                        slide_dict["content"].append(
                                            " " * paragraph.level + "- " + paragraph.text)
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
            #elif shape.shape_type == 13:
                # picture
                # get the picture from the slide
            #    picture = shape.image.blob
                # save the picture blob to the dictionary
            #    slide_dict["pictures"].append(picture)
            else:
                continue


        except AttributeError:
            continue
    return slide_dict

def get_pptx_presentation(file):
    subject_items = file.split("\\")[5].split(" ",1)
    presentation_dict = {
        "title": "",
        "slides": []
    }
    try:
        prs = pptx.Presentation(file)
        pptx_slides = get_pptx_slides(prs)
        title_slide = pptx_slides[0]
        if (title_slide["title"] == subject_items[0] or title_slide["title"] == subject_items[-1]) or title_slide["title"] == "":
            try:
                title_slide["title"] = title_slide["content"][0]
                title_slide["content"] = title_slide["content"][1:]
            except IndexError:
                pass
        presentation_dict["title"] = title_slide["title"]
        presentation_dict["slides"] = pptx_slides[1:]
        return presentation_dict
    except pptx.exc.PackageNotFoundError:
        return None
class ObsidianVault:
    def __init__(self, path):
        self.path = path

    def create_folder(self, folder_path):
        if not os.path.exists(os.path.join(self.path, folder_path)):
            os.mkdir(os.path.join(self.path, folder_path))
        else:
            pass

    def create_file(self, file_path, content):
        with open(os.path.join(self.path, file_path), "w") as file:
            file.write(content)

vault = ObsidianVault(r"C:\Users\Troy\Obsidian\College")



def main():
    for file in files:
        extension = file.split(".")[-1]
        if extension == "pptx":
            pptx_file = get_pptx_presentation(file)
            #print(pptx_file)
if __name__ == "__main__":
    main()