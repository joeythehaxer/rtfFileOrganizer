import os

FOLDER_PATH_PROMPT = "Enter the folder path: "
EMPTY_SUBFOLDERS_FOUND_MESSAGE = "Empty subfolders:"
NO_EMPTY_SUBFOLDERS_FOUND_MESSAGE = "No empty subfolders found."


def get_empty_subfolders(folder):
    empty_subfolders = []
    for root, dirs, files in os.walk(folder):
        for d in dirs:
            dir_path = os.path.join(root, d)
            if not os.listdir(dir_path):
                empty_subfolders.append(dir_path)
    return empty_subfolders


def print_empty_subfolders(empty_subfolders):
    if empty_subfolders:
        print(EMPTY_SUBFOLDERS_FOUND_MESSAGE)
        for subfolder in empty_subfolders:
            print(subfolder)
        print(f"Total empty subfolders: {len(empty_subfolders)}")
    else:
        print(NO_EMPTY_SUBFOLDERS_FOUND_MESSAGE)


def main():
    folder_path = input(FOLDER_PATH_PROMPT)
    empty_subfolders = get_empty_subfolders(folder_path)
    print_empty_subfolders(empty_subfolders)


if __name__ == "__main__":
    main()
