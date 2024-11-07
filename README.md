# Vscode vba editing tool
Vscode tool with python and [xlwings library cli](https://docs.xlwings.org/en/latest/command_line.html) to easily edit and sync VBA code from a repository to excel. Based on this [xlwings tutorial](https://youtu.be/xoO-Fx0fTpM)

## How to use?
### 1. Download
Make sure [python](https://www.python.org/) is installed. Then [clone](https://git-scm.com/docs/git-clone) this repository to your machine using your preffered method.

### 2. Configure
Open the repo on [Vscode](https://code.visualstudio.com/) or [Cursor](https://www.cursor.com/). After that you can press `ctrl + shif + B` and select the `Configure Environment` vscode task, this is a one time action and you won't need to execute this anymore.

### 3. Usage (Vscode tasks)
To use this tool you need to be on Vscode or Cursor IDE's, the whole project it's based on Vscode tasks so it won't work on other IDE's.
By pressing `ctrl + shif + B` some options will appear:
1. Configure Environment
2. Create new project
3. Create file
4. Vba EDIT
5. Vba IMPORT

#### 3.1 Configure Environment
This task will configure the virtual environment (`.venv`) for the give python version you're using and install on this `.venv` [requirements](requirements.txt) needed for the project.

#### 3.2 Create a new project
This task will create on this repo a folder based on our [model](./.docs/model) to concentrate all your VBA `.bas` files for a given `.xlsm`, and if you're using git by forking this repo, or another method, you have the advantage of version control.
<br>A project consists of a folder with VBA files on it and a [metadata.yml](./.docs/model/metadata.yml) that contains useful information about the project or specific files, configure it as you wish by now, but in the future this will be standardized with a custom built python manager.
<br>My suggestion is, create one project for each `.xlsm` file as VBA Import will import everything.

#### 3.3 Create file
This task will create on a folder a `.bas` file based on this [exemple.bas](./.docs/model/exemple.bas) file, it's important to use this task or follow the recommendations in [4.1](README.md#41-file-creation).This task will prompt you for:
* What's the path where the VBA `.bas` will be stored.

#### 3.4 Vba EDIT
This task manages to open the `.xlsm` file and export all the existing VBA `.bas` on it, this is a destructive action so if you have other `.bas` files with the same name, they'll be subscripted. This task will prompt you for:
* What's the path where the VBA `.bas` will be stored.
* What's the path for the `.xlsm` file that will be edited.

#### 3.5 Vba IMPORT
This task manages to open the `.xlsm` file and import all the existing VBA `.bas` on a given directory. This task will prompt you for:
* What's the path where the VBA `.bas` are stored.
* What's the path for the `.xlsm` file that will be edited.

### 4 Important considerations
#### 4.1 File creation
When creating new VBA `.bas` file, use the `Create file` task or place at the top of the file the below code __it's extremely importatnt__:
```
Attribute VB_Name = "__file__"
```
That way xlwings can know which name to add onto the `.xlsm` file, or you can use the `Create file` task.

#### 4.2 System support
The current project was created and tested on python 3.12 on a Windowns 11 machine, currently MacOs is not supported, but it'll be soon.

### 5 Troubleshooting
* xlwings errors: Verify if macros are enabled, see [this link](https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-microsoft-365-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6)
* this error `pywintypes.com_error: (-2147352567, 'Exceção.', (0, 'Microsoft Excel', 'O acesso de programação ao projeto do Visual Basic não é confiável\n', 'xlmain11.chm', 0, -2146827284), None)`, see [this link](https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-microsoft-365-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6)
* shell errors: Verify if python is installed, the `.venv` setted and [requirements](./requirements.txt) installed.

### 6. Extra tips
* [VBA Best practices](https://www.automateexcel.com/vba/best-practices/)
* [Maintanable VBA](https://masterofficevba.com/vba-coding-constructs/maintainable-excel-vba-code-best-practices/)
* [VBA Best practices (Advanced)](https://masterofficevba.com/vba-coding-constructs/best-practices-for-excel-vba-code/)
