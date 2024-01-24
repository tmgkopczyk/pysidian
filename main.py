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
        file_details = file.split("\\")[4:]
        semester = file_details[0]
        course = file_details[1]
        file_name = file_details[-1]
        file_extension = file_name.split(".")[-1]


if __name__ == '__main__':
    main()
