"""
Collection of the more genal functions used in the project
"""

import os
import pathlib


def get_files_from_folder(folder_path):
    """Gets list of files in provided folder path"""
    folder = pathlib.Path(folder_path)
    return list(folder.glob("*.xls"))