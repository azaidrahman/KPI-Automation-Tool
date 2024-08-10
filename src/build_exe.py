import os
import subprocess
import shutil
import sys

def build_exe():
    # Ensure PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("PyInstaller is not installed. Installing now...")
        subprocess.call([sys.executable, "-m", "pip", "install", "pyinstaller"])

    # Set the name of your main script and the desired exe name
    main_script = "main.py"
    exe_name = "pickwick_CSM"

    # Set the icon file path (adjust as needed)
    icon_path = "../media/pickwick_icon.ico"  # Replace with your icon path if you have one

    # Create the PyInstaller command
    command = [
        "pyinstaller",
        "--name={}".format(exe_name),
        "--onefile",
        "--windowed",
        "--add-data={}".format(os.path.join("src", "*:src")),
    ]

    # Add icon if the file exists
    if os.path.exists(icon_path):
        command.append("--icon={}".format(icon_path))

    command.append(main_script)

    # Run PyInstaller
    subprocess.call(command)

    # Move the executable to the project root
    dist_dir = "dist"
    if os.path.exists(dist_dir):
        for file in os.listdir(dist_dir):
            shutil.move(os.path.join(dist_dir, file), ".")
        os.rmdir(dist_dir)

    # Clean up build directory and spec file
    build_dir = "build"
    spec_file = "{}.spec".format(exe_name)
    if os.path.exists(build_dir):
        shutil.rmtree(build_dir)
    if os.path.exists(spec_file):
        os.remove(spec_file)

    print("Build complete. Executable created: {}.exe".format(exe_name))

if __name__ == "__main__":
    build_exe()