# Project Description

The project aims to create a search tool that allows users to find, execute, and enable or disable scheduled tasks on a server from a local environment. It provides a user-friendly interface for managing scheduled tasks remotely without directly accessing the server.


# Installation Guide

This guide will help you set up the necessary environment to run the project. Follow the steps below for a successful installation.

## Prerequisites

- Python: Make sure you have Python installed on your computer. You can download it from [python.org](https://www.python.org). Be sure to add Python to the PATH during the installation.

## Installation

1. Download the project from GitHub using the "Clone or download" button on the repository page.

2. Run the `Instalabibliotecas.bat` file to install the required libraries for the project. This file is located in the downloaded project folder.

3. Open the provided `.ini` file and modify its contents as follows:

   ```ini
   [Credentials]
   server = "your_server"
   user = "your_user" (user with admin power in the server)
   password = "your_password" (password to that user)
   
   
  Replace "your_server", "your_username", and "your_password" with the appropriate values. Save the file after making the changes.
  

4. Run the Encryptor.ini file to encrypt the modified .ini file. This process will generate a password. Take note of this password.
   
   `C:\\Path\to\file encryptor.py`

6. Copy that password into the key.txt file 

## Execution

Once you have completed the installation, you are ready to run the project. You can use one of the following files:

- `Buscatareas.bat` : Run this file to start the project from the command line.

- `Buscatareas.pyw` : Run this file to start the project in a windowed environment.

Both files perform the same function, so you can choose whichever suits your needs.

That's it! Now you can enjoy your project. If you have any questions or encounter any issues during the installation, feel free to contact me.
