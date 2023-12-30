import os
import pptx

# create a Python function to get a list of files in a directory, including all files within subdirectories and return them as full paths
def get_filepaths(directory):
    # list of file paths
    file_paths = []
    # walk the directory tree
    for root, directories, files in os.walk(directory):
        for filename in files:
            if "Slides" not in root:
                continue
            else:
                # join the two strings to form the full filepath
                filepath = os.path.join(root, filename)
                file_paths.append(filepath)
    return file_paths

def get_file_extensions(files):
    file_extensions = []
    for i in range(len(files)):
        file = files[i]
        file_data = {"subject": file.split("\\")[5], "filename": file.split("\\")[-1][:file.index(".") - (len(file))],
                     "file_extension": file.split(".")[-1],"term":file.split("\\")[4],"file_path":file}
        if "Numeracy and Logic" in file_data["subject"]:
            file_data["module"] = file.split("\\")[-2]
        file_extensions.append(file_data)
    return file_extensions

def organize_files(file_list):
    organized_files = {}
    # organize files by term, and then by subect
    for file in file_list:
        if file["term"] not in organized_files.keys():
            organized_files[file["term"]] = {}
        if file["subject"] not in organized_files[file["term"]].keys():
            organized_files[file["term"]][file["subject"]] = []
        organized_files[file["term"]][file["subject"]].append(file)
    return organized_files

def convert_networking_fundamentals_presentation(file):
    presentation = {}
    if file.endswith("pptx"):
        prs = pptx.Presentation(file)
        # get the presentation title from the title of the first slide
        try:
            presentation["title"] = prs.slides[0].shapes.title.text
        except AttributeError:
            presentation["title"] = "Untitled"
        # get the presentation slides
        slides = get_networking_slides(prs.slides)
        presentation["slides"] = slides
    return presentation

def get_networking_slides(slides):
    slide_list = []
    section_dict = {}
    section_slides = []
    for slide_index,slides in enumerate(slides):
        if slide_index == 0:
            continue
        else:
            slide_list.append(get_networking_slide_content(slides))
    return slide_list

def get_networking_slide_content(slide):
    slide_dict = {
        "title": "",
        "content": []
    }
    try:
        slide_dict["title"] = slide.shapes.title.text
    except AttributeError:
        slide_dict["title"] = "Untitled"
    for shape in slide.shapes:
        if shape.has_text_frame:
            if shape.text_frame.text.strip() != "" and shape.text_frame.text.strip() != slide_dict["title"]:
                if len(shape.text_frame.text.strip()) > 2:
                    slide_dict["content"].append(shape.text_frame.text.strip())
                else:
                    continue
            else:
                continue
    return slide_dict

def main():
    files = get_filepaths(r"C:\Users\Troy\Algonquin\Fall 2023")
    file_extensions = get_file_extensions(files)
    organized_files = organize_files(file_extensions)
    for term in organized_files.keys():
        for subject in organized_files[term].keys():
            for file in organized_files[term][subject]:
                if file["subject"] == "Networking Fundamentals":
                    presentation = convert_networking_fundamentals_presentation(file["file_path"])
                    print(presentation)


if __name__ == '__main__':
    main()