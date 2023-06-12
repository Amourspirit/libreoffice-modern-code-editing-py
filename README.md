# LibreOffice Python Script Modern Code Editor Examples

## Overview

This repository demonstrates how to use a modern IDE to edit LibreOffice Python scripts with Type support, debugging and Testing.

### Prerequisites

- Basic knowledge of python
- Basic knowledge of using a terminal
- IDE installed such as Vs Code
- Basic knowledge using python macros in LibreOffice

## Set up Environment

This repository is set to use a pip environment.

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

### Windows

Python environments for LibreOffice on Windows need to be configured to use LibreOffice Embedded python.
One of the drawbacks of this is LibreOffice Embedded python has limitations.
See the guide [Manually Creating a Virtual Environment](https://python-ooo-dev-tools.readthedocs.io/en/latest/guide/virtual_env/windows_manual_venv.html) for more information.

The recommended way to set up a python environment for LibreOffice on windows is to use a [Pre-configured virtual environment](https://github.com/Amourspirit/lo-support_file/tree/main/virtual_environments/windows).
If there is a pre-configured environment that matches your LibreOffice Embedded Python version then the short version of getting started is to extract the pre-configured zip into your current project root and then after activating pip install requirements.txt (`pip install -r requirements.txt`).

If a pre-configured environment does not work for you then you can follow  the guide [Manually Creating a Virtual Environment](https://python-ooo-dev-tools.readthedocs.io/en/latest/guide/virtual_env/windows_manual_venv.html) .

After the virtual environment is created then the next steps are to activate and install requirements.

#### Activate

Activate with PowerShell:

```powershell
.\.venv\Scripts\activate.ps1
```

Activate with CMD

```shell
.\.venv\Scripts\activate.bat
```

#### Install Requirements

It is important that pip be called as a python module in this case.

```powershell
python -m pip install -r requirements.txt
```

Installation may take a bit of time.

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

more to come ...
