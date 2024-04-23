import os.path
from zipfile import ZipFile, ZIP_DEFLATED


def unzip_file(zip_file_path: str, save_path: str):
    with ZipFile(zip_file_path, 'r') as zip_obj:
        zip_obj.extractall(save_path)


def zip_file(target_file_path, save_path: str):
    with ZipFile(save_path, "w", ZIP_DEFLATED) as zip_obj:
        for root, dirs, files in os.walk(target_file_path):
            for directory in dirs:
                directory_path = os.path.join(root, directory)
                relative_path = os.path.relpath(directory_path, target_file_path)
                zip_obj.write(directory_path, relative_path)
            for file in files:
                if file == ".DS_Store":
                    continue
                full_path = os.path.join(root, file)
                relative_path = os.path.relpath(full_path, target_file_path)
                zip_obj.write(full_path, relative_path)


if __name__ == '__main__':
    # unzip_file('/Users/grantit/Desktop/人工智能.pptx', '/Users/grantit/Desktop/AI_test')
    zip_file('/Users/grantit/Desktop/AI_test', '/Users/grantit/Desktop/AI_test.pptx')
