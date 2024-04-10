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
        if not slide_content.get("content"):
            continue
        else:
            slides.append(slide_content)
    return slides


def get_pptx_slide_content(slide):
    slide_dict = {
        "title": "",
        "content": [],
        "pictures": []
    }
    # get the title of the slide
    image_count = 0
    for shape in slide.shapes:
        try:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    if len(paragraph.text.strip()) <= 2:
                        continue
                    else:
                        # append paragraph text to slide content based on level
                        if paragraph.level == 0:
                            slide_dict["content"].append(paragraph.text.strip())
                        elif paragraph.level > 0:
                            slide_dict["content"].append(
                                " " * paragraph.level + "- " + paragraph.text.strip())
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
            if shape.shape_type == 13:
                picture = shape.image
                slide_dict["pictures"].append(picture.blob)
            else:
                continue

        except AttributeError:
            continue
    if slide_dict.get("content"):
        slide_title = slide_dict["content"][0]
        x0b_match = re.findall(r"[\x0b\xa0]", slide_title)
        if len(x0b_match) == 0:
            slide_dict["title"] = slide_title
            slide_dict["content"] = slide_dict["content"][1:]
        if len(x0b_match) == 1:
            _title_list = re.split("\x0b", slide_title)
            if _title_list[0][-1] == " " or _title_list[0][-1] == "," or _title_list[0][-1] == ":":
                _title_list = " ".join([x.strip() for x in _title_list])
            else:
                slide_dict["title"] = _title_list[-1]
                slide_dict["section"] = _title_list[0]
                slide_dict["content"] = slide_dict["content"][1:]
        elif len(x0b_match) >= 2:
            _title_list = " ".join(x for x in list(filter(None,re.split("\x0b", slide_title))) if x.strip() !="")
            slide_dict["title"] = _title_list
            slide_dict["content"] = slide_dict["content"][1:]
    return slide_dict


def get_pptx_presentation(file):
    subject = file.split("\\")[5]
    subject_items = subject.split(" ", 1)
    presentation_dict = {
        "title": "",
        "slides": [],
        "subject": subject
    }
    try:
        prs = pptx.Presentation(file)
        pptx_slides = get_pptx_slides(prs)
        title_slide = pptx_slides[0]
        if subject.startswith("CST8182") or subject.startswith("CST8315"):
            if title_slide["title"] == "Introduction to Networks v7.0 (ITN)" or title_slide["title"] == "Switching, Routing, and Wireless Essentials v7.0 (SRWE)":
                title_slide["content"].append(title_slide["title"])
                title_slide["title"] = title_slide["content"][0]
        if (title_slide["title"] == subject_items[0] or title_slide["title"] == subject_items[-1]) or title_slide[
            "title"] == "":
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
        try:
            with open(os.path.join(self.path, file_path), "w", encoding="utf-8") as file:
                file.write(content)
        except OSError:
            pass


vault = ObsidianVault(r"C:\Users\Troy\Documents\Obsidian\College Notes")


def main():
    for file in files:
        extension = file.split(".")[-1]
        if extension == "pptx":
            pptx_file = get_pptx_presentation(file)
            create_presentation_directory_structure(pptx_file)


def create_presentation_directory_structure(presentation):
    invalid_chars = re.compile(r"[\\/:*?<>|]")
    presentation_title = re.sub(r"[\\/:*?<>|]","",presentation['title'])
    if not os.path.exists(os.path.join(vault.path,presentation['subject'])):
        os.mkdir(os.path.join(vault.path,presentation['subject']))
    else:
        pass
    presentation_folder = str(os.path.join(vault.path, presentation['subject'], presentation_title))
    if not os.path.exists(presentation_folder):
        os.mkdir(presentation_folder)
    else:
        pass
    for slide in presentation['slides']:
        slide_title = re.sub(r"[\\/:*?<>|]","",slide['title'])
        if slide.get("section"):
            slide_section = re.sub(r"[\\/:*?<>|]","", slide["section"])
            section_path = os.path.join(presentation_folder,slide_section)
            if not os.path.exists(section_path):
                os.mkdir(section_path)
            slide_path = os.path.join(section_path,slide_title)
            try:
                with open(f"{slide_path}.md","w",encoding="utf-8") as f:
                    for content in slide["content"]:
                        f.write(content)
                        f.write("\n")
            except OSError:
                pass
            if slide.get("pictures"):
                for p_index, picture in enumerate(slide["pictures"]):
                    try:
                        with open(f"{slide_path}_{p_index}.png","wb") as image:
                            image.write(picture)
                        with open(f"{slide_path}.md","a",encoding="utf-8") as f:
                            f.write("\n")
                            f.write(f"![[{slide_title}_{p_index}.png]]")
                    except OSError:
                        pass

        else:
            slide_path = os.path.join(presentation_folder,slide_title)
            try:
                with open(f"{slide_path}.md","w",encoding="utf-8") as f:
                    for content in slide["content"]:
                        f.write(content)
                        f.write("\n")
            except OSError:
                pass
            if slide.get("pictures"):
                for p_index, picture in enumerate(slide["pictures"]):
                    try:
                        with open(f"{slide_path}_{p_index}.png", "wb") as image:
                            image.write(picture)
                        with open(f"{slide_path}.md", "a", encoding="utf-8") as f:
                            f.write("\n")
                            f.write(f"![[{slide_title}_{p_index}.png]]")
                    except OSError:
                        pass


if __name__ == "__main__":
    main()
