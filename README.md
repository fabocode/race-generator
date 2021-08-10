# Race Generator

A Python program to take user inputs to generate waypoints in pulses/gps for racing applications


# Installation

It's required to run the following command to install a prerequisite `openpyxl` module

    $ pip install -r requirements.txt

# Usage

To use it:

    $ python cli.py

or if you use python3:

    $ python3 cli.py

# Tip to Generate an executable version

Call `pyinstaller` with `--paths` argument to add the installed modules in the virtual environment on windows 10:

    $ pyinstaller --onefile --name race_generator --paths=C:\Users\{your_python_path}\env\Lib\site-packages cli.py