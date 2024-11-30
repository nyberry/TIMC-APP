import os, sys
from cx_Freeze import setup, Executable

# may need to chdir to something like: C:\Users\nckbr\Desktop\TIMC_APP - python 4 Feb 2024cddir
current_directory = os.getcwd()
if current_directory.endswith('src'):
    new_directory = current_directory[:-4]  # Remove the last four characters ('\src')
    os.chdir(new_directory)
    print(f"Current working directory changed to: {os.getcwd()}")
else:
    print(f"Current working directory: {os.getcwd()}")


# Specify the script you want to compile
script = "src/main.py"

# Specify the output directory for the .exe file
output_directory = "C:/Users/nckbr/Desktop/TIMC_APP"

base = None
if sys.platform == "win32":
    base = "Win32GUI"  # Use this for a GUI application on Windows

# Options for the executable
build_options = {
    "packages": [],  # List of packages to include
    "excludes": [],  # List of packages to exclude
    "include_files": ['data','images','temp','src','install','install.bat','copy2USB.bat'],  # List of additional files to include
}

# Create the Executable object
executable = Executable(
    script=script,
    base = "Win32GUI",
    target_name="TIMC_APP.exe",  # Output executable file name
)

# Setup parameters
setup(
    name="TIMC_APP",
    version="2.1",
    description="Make patient reports at TIMC",
    options={"build_exe": build_options},
    executables=[executable],
    # Specify the output directory for the executable
    script_args=["build", "--build-exe=" + output_directory],
)
