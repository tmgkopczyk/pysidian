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

files = get_filepaths(r"C:\Users\Troy\Algonquin\Fall 2023")

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

file_extensions = get_file_extensions(files)

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

organized_files = organize_files(file_extensions)

def make_obsidian_folders(obsidian_directory,organized_folders):
    for term in organized_folders.keys():
        term_directory = os.path.join(obsidian_directory,term)
        print(term_directory)
        for subject in organized_folders[term].keys():
            subject_directory = os.path.join(term_directory,subject)
            print(subject_directory)
            for file in organized_files[term][subject]:
                if file.get("module") != None:
                    module_directory = os.path.join(subject_directory,file["module"])
                    print(module_directory)
                else:
                    print(subject_directory)

make_obsidian_folders(r"C:\Users\Troy\Obsidian\College",organized_files)