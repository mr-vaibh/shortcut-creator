import os
import time
import pylnk3
import sys

GIT_FOLDER = r"D:\~GIT"
SHORTCUT_FOLDER = r"D:\Shortcuts"
CHECK_INTERVAL = 3600  # 1 hour


def create_shortcut(target_folder, shortcut_path):
    """Creates a Windows shortcut for a folder."""
    lnk = pylnk3.create(shortcut_path, target=target_folder)
    lnk.write()


def update_shortcuts():
    """Updates shortcuts based on the current state of D:\~GIT."""
    if not os.path.exists(SHORTCUT_FOLDER):
        os.makedirs(SHORTCUT_FOLDER)

    existing_shortcuts = {f for f in os.listdir(SHORTCUT_FOLDER) if f.endswith('.lnk')}
    current_folders = {f for f in os.listdir(GIT_FOLDER) if os.path.isdir(os.path.join(GIT_FOLDER, f))}

    # Create or update shortcuts
    for folder in current_folders:
        target_path = os.path.join(GIT_FOLDER, folder)
        shortcut_path = os.path.join(SHORTCUT_FOLDER, f"{folder}.lnk")

        if folder + ".lnk" not in existing_shortcuts:
            create_shortcut(target_path, shortcut_path)
        existing_shortcuts.discard(folder + ".lnk")  # Remove from deletion list

    # Remove obsolete shortcuts
    for obsolete in existing_shortcuts:
        os.remove(os.path.join(SHORTCUT_FOLDER, obsolete))


def main():
    while True:
        update_shortcuts()
        time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    # Run the script with pythonw.exe to hide the console
    main()
