import os


def del_all_files(path):
    son_folders = []
    for file in os.listdir(path):
        file_path = os.path.join(path, file)
        if os.path.isdir(file_path):
            son_folders.append(file_path)
            del_all_files(file_path)
        else:
            os.remove(file_path)
    for it in reversed(son_folders):
        os.rmdir(it)
    os.rmdir(path)


if __name__ == '__main__':
    path = "/Users/grantit/Desktop/人工智能/ppt/charts/chart1"
    print(os.path.splitext(path))