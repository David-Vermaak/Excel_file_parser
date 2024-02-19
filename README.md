# Excel_file_parser
This is an Excel file parser written in Python 

I am using the sun valley ttk theme from https://github.com/rdbende/Sun-Valley-ttk-theme

Feel free to use this project to build nice looking python programs!

## Compiling the .exe

make sure that pyinstaller is installed:

```bash
pip install pyinstaller
```

Then run this command in the termainal (within the project folder)
Make sure that the path of the sv_ttk dependency is correct. (replace USERNAME)

```bash
pyinstaller -w -F -i "email.ico" --add-data "C:\Users\USERNAME\Desktop\eft-emails\sv_ttk;sv_ttk" Excel_file_parser.py
```


![app](https://github.com/David-Vermaak/Excel_file_parser/assets/100315563/507bfe09-b3e2-4a38-84ca-71a869ccf60d)

![lightapp](https://github.com/David-Vermaak/Excel_file_parser/assets/100315563/d86920c3-6f79-4dcf-983a-26dfe89a5c80)

![excel](https://github.com/David-Vermaak/Excel_file_parser/assets/100315563/e235b8e4-6475-45fa-83ff-a8a645238cc3)

![excelLight](https://github.com/David-Vermaak/Excel_file_parser/assets/100315563/b0dcb019-b1b9-4a80-97c1-a652d8a3afda)
