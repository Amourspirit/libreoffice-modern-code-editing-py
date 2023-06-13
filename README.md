# LibreOffice Python Script Modern Code Editor Examples


- [LibreOffice Python Script Modern Code Editor Examples](#libreoffice-python-script-modern-code-editor-examples)
  - [Overview](#overview)
    - [Prerequisites](#prerequisites)
  - [Set up Environment](#set-up-environment)
  - [Linux/Mac](#linuxmac)
    - [Prerequisites](#prerequisites-1)
    - [Create Virtual Environment](#create-virtual-environment)
    - [Activate Virtual Environment](#activate-virtual-environment)
      - [Install Requirements](#install-requirements)
      - [Install local project](#install-local-project)
  - [Windows](#windows)
    - [Create Virtual Environment](#create-virtual-environment-1)
    - [Activate](#activate)
    - [Install Requirements](#install-requirements-1)
    - [Install local project](#install-local-project-1)
  - [Test Environment](#test-environment)
  - [Open in Code Editor](#open-in-code-editor)
    - [Vs Code](#vs-code)
      - [Install Profile](#install-profile)
      - [Select Interpreter](#select-interpreter)
        - [Windows](#windows-1)
      - [Linux/Mac](#linuxmac-1)
      - [Running Scripts](#running-scripts)
      - [Running Test in Vs Code](#running-test-in-vs-code)


## Overview

This repository demonstrates how to use a modern IDE to edit LibreOffice Python scripts with Type support, debugging and Testing.

### Prerequisites

- Basic knowledge of python
- Basic knowledge of using a terminal
- IDE installed such as Vs Code
- Basic knowledge using python macros in LibreOffice

## Set up Environment

This repository is configured to use virtual environment and [pip](https://pip.pypa.io/en/stable/) as a package manager.

See [Virtual Environments Guides](https://python-ooo-dev-tools.readthedocs.io/en/latest/guide/virtual_env/index.html) that are specifically written for creating Virtual Environments for LibreOffice.

## Linux/Mac

There is a guide, [Manually Creating a Virtual Environment](https://python-ooo-dev-tools.readthedocs.io/en/latest/guide/virtual_env/linux_manual_venv.html) that provides the instructions for setting up a virtual environment in Linux.
It is recommended to follow this guide, so we will summarize here.

### Prerequisites

The `libreoffice-script-provider-python` apt package must be installed. This package allows scripts to connect to LibreOffice.

```bash
sudo apt install libreoffice-script-provider-python
```

### Create Virtual Environment

In your project folder run the following command.

```bash
/usr/bin/python3 -m venv .venv
```

### Activate Virtual Environment

Activate the virtual environment.

```bash
source .venv/bin/activate
```

#### Install Requirements

It is important that pip be called as a python module in this case.

```bash
python -m pip install -r requirements.txt
```

Installation may take a bit of time.

Set virtual environment to work with LibreOffice.

```bash
oooenv cmd-link -a
```

#### Install local project

Install local project into virtual environment. This makes it so the test can see the project files when run on the command line.
Similar option as `pip -e .`

```bash
oooenv -e
```

## Windows

Python environments for LibreOffice on Windows need to be configured to use LibreOffice Embedded python.
One of the drawbacks of this is LibreOffice Embedded python has limitations.
See the guide [Manually Creating a Virtual Environment](https://python-ooo-dev-tools.readthedocs.io/en/latest/guide/virtual_env/windows_manual_venv.html) for more information.

### Create Virtual Environment

The recommended way to set up a python environment for LibreOffice on windows is to use a [Pre-configured virtual environment](https://github.com/Amourspirit/lo-support_file/tree/main/virtual_environments/windows).
If there is a pre-configured environment that matches your LibreOffice Embedded Python version then the short version of getting started is to extract the pre-configured zip into your current project root and then after activating pip install requirements.txt (`pip install -r requirements.txt`).

If a pre-configured environment does not work for you then you can follow  the guide [Manually Creating a Virtual Environment](https://python-ooo-dev-tools.readthedocs.io/en/latest/guide/virtual_env/windows_manual_venv.html) .

After the virtual environment is created then the next steps are to activate and install requirements.

### Activate

Activate with PowerShell:

```powershell
.\.venv\Scripts\activate.ps1
```

Activate with CMD

```shell
.\.venv\Scripts\activate.bat
```

### Install Requirements

It is important that pip be called as a python module in this case.

```powershell
python -m pip install -r requirements.txt
```

Installation may take a bit of time.

### Install local project

Install local project into virtual environment. This makes it so the test can see the project files when run on the command line.
Similar option as `pip -e .`

```powershell
oooenv -e
```

## Test Environment

At this point you should be able to run python and import `uno` for a test.

Withe the Virtual Environment active, run the command `$ python`, this will put you in the python console.
Import `uno` if there are no import errors then you virtual environment should be configured correctly.

```python
>>> import uno
>>> 
```

To exit Python console just type `exit()` and press the enter key.

```python
>>> exit()
```

## Open in Code Editor

### Vs Code

Start Vs Code:

From the command prompt you of your project you can type `code .` to start Vs Code in the current directory.
Alternatively you can start Vs Code normally and open the folder where this repo is stored.

#### Install Profile

If you are familiar with Vs Code and you already have Vs Code set up to use with Python then this profile is probably not needed.
However if you are new to Vs Code then this is a fast way to set the Vs Code extensions to run a minimal number of extensions to get you up and running to use these examples.

See [Profiles in Visual Studio Code](https://code.visualstudio.com/docs/editor/profiles).

The following link contains a profile configuration that contains all the extensions needed to run these examples.

```text
https://gist.github.com/Amourspirit/b44186bc6838702ffb1b21d2e6100748
```

Open Vs Code command pallet `shift+ctl+p` and type in `import profile`.
Choose `Import Profile...`

![2023-06-13_11-06-25](https://github.com/Amourspirit/libreoffice-modern-code-editing-py/assets/4193389/0df7dd3b-6eaf-4ff8-b7a6-23ee87256987)

Enter in the link above and press enter.

![2023-06-13_11-05-29](https://github.com/Amourspirit/libreoffice-modern-code-editing-py/assets/4193389/5ff090e6-b2f0-4da4-828f-b6e19c55460f)

#### Select Interpreter

When the project is firs loaded into Vs Code you will need to select a Python Interpreter.
The interpreter is the python used to run the python examples in this project.
We will need to select the python that was setup in the `.venv` directory.

One way to accomplish this is to open the command pallet `shift+ctl+p` and type in `select interpreter`

![2023-06-13_11-23-57](https://github.com/Amourspirit/libreoffice-modern-code-editing-py/assets/4193389/d9529073-1960-473a-98ff-ead5912b1d56)

##### Windows

The interpreter for windows should be:

```powershell
.\.venv\Scripts\python.exe
```

![2023-06-13_11-21-35](https://github.com/Amourspirit/libreoffice-modern-code-editing-py/assets/4193389/26a1a633-69df-4d71-85d6-aea0f4c9dfac)

**Note**: For this setup make sure the interpreter ends with `.exe`

#### Linux/Mac

```bash
./.venv/bin/python
```

![2023-06-13_11-52-33](https://github.com/Amourspirit/libreoffice-modern-code-editing-py/assets/4193389/eea5b606-fe6a-46cc-b9b9-91440ed5c207)

#### Running Scripts

The scripts in this project can be run as stand alone scripts, or the scripts can be embedded into a document as well.

With the help of [OOO Development Tools](https://python-ooo-dev-tools.readthedocs.io/en/latest/index.html) (OooDev) Starting LibreOffice automatically and Creating or Opening documents is somewhat straight forward.

The `src/hello_world/hello_world.py` is intended to be a simple example. When Run from python or Vs Code it Opens Writer, creates a new document, writes "Hello World", and then closes Writer.

The `src/lib_o_con_2021/lib_o_con_2021_calc.py` has a bit of history behind it.
Originally the script was used in the talk "Python scripts in LibreOffice Calc using the [ScriptForge](https://gitlab.com/LibreOfficiant/scriptforge) library" given by Rafael Lima during the LibreOffice Conference 2021. Original on [github](https://github.com/rafaelhlima/LibOCon_2021_SFCalc). Then the [LibreOffice Python UNO Examples](https://github.com/Amourspirit/python-ooouno-ex) project extended it with type support which can be seen [here](https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/calc/lib_o_con_2021). Now it is extended and used in this project as an example.

The results of running the `src/lib_o_con_2021/lib_o_con_2021_calc.py` can be see below.
The script connects to LibreOffice and runs the methods that are active in the `_debug_methods()` method, Then LibreOffice is closed.

<https://github.com/Amourspirit/libreoffice-modern-code-editing-py/assets/4193389/6f271e21-34a8-420c-a901-53283458157c>

#### Running Test in Vs Code

**Note** this project is set up to up to use [pytest](https://docs.pytest.org/), but it could just as easily use built in python unittest.

In the testing Section of Vs Code you can bring up the available test.
![2023-06-13_13-44-09](https://github.com/Amourspirit/libreoffice-modern-code-editing-py/assets/4193389/2b3952b3-f053-4d0f-8577-c2a997e6e126)

As seen here you can debug the test like any other Python Test.

<https://github.com/Amourspirit/libreoffice-modern-code-editing-py/assets/4193389/8ae434e6-496d-40d9-a2e1-062e05d1bffd>
