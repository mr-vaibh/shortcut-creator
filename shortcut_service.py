import os
import time
import winshell
import pythoncom

GIT_FOLDER = r"D:\~GIT"
SHORTCUT_FOLDER = r"D:\Shortcuts"
CHECK_INTERVAL = 3600  # 1 hour (3600 seconds)


def create_shortcut(target_folder, shortcut_path):
    """Creates a Windows shortcut for a folder using winshell."""
    with winshell.shortcut(shortcut_path) as link:
        link.path = target_folder
        link.description = f"Shortcut to {os.path.basename(target_folder)}"


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
    pythoncom.CoInitialize()  # Required for COM operations
    while True:
        update_shortcuts()
        time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    main()
