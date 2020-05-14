# option-sheets-rec
## To Edit Source Code and Work with GIT
1. Use Git Bash
2. `cd ../../Development`
3. `git clone git@github.com:northriverboats/option-sheets-rec.git`
4. `cd option-sheets-rec`

5. Use windows shell
6. `cd \Development\option-sheets-rec`
7. `\Python37\python -m venv venv`
8. `venv\Scripts\activate`
9. `python -m pip install pip --upgrade`
10. `pip install -r requirements.txt`
11. Remember to Create New Branch Before Doing Any Work

## Build Executable
`venv\Scripts\pyinstaller.exe --onefile --windowed --icon options.ico  --name "Option Sheets for Recreational" "option-sheet-rec.spec" main.py`