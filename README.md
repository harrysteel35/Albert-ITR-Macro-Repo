# Albert's Magic ITR Spreadsheet
![Welcome image](https://tenor.com/en-GB/view/welcome-to-the-fucking-show-welcome-gif-8992631.gif)



## Instructions for Users
1. Click the big green button above that says **\"<> Code\"**
2. Click **\"Download ZIP\"**
   * Note that this should just download to your \"Downloads\" folder on your computer
3. Extract the contents of the .zip folder
   * In the folder where the .zip file is saved, right click on it, then click
 **"Extract All...\"**
   * After this a window should pop up. Click \"**Extract**\".
   * The files will be extracted to a folder with the same name as the .zip, copy these into a folder of your choosing.
4. Add the macro to Excel
   * Open \"ITR_Template.xlsm\"
   * If there is a warning saying \"Macros have been disabled\", click the **Enable** button
   * In Excel, open the **Developer** tab
   * Click **Visual Basic** and a new window will pop up, titled *Microsoft Visual Basic for Applications*
   * In the **Project** tab of this window, right click on **\"VBAProject (ITR_Template.xlsm)\"**
   * Click **\"Import File...\"**
   * Navigate through your file system to where you saved the contents of this repository
   * Add all .bas and .frm files. You may need to do these one by one.
   * Once this is done, close the *Microsoft Visual Basic for Applications* window
5. Running the macro/application
   * In the **Developer** tab of Excel, click **Macros**
   * Click \[Insert macro name here\]
   * Click Run
   * Follow the instructions and prompts




## Instructions for Devs
### Installing Git Bash
* Download Git Bash from https://git-scm.com/downloads
* Install
  

### Cloning this repository
* Note that Git Bash is basically a linux terminal. The following commands may be useful for navigating the file system:
  - **pwd** = "Print working directory": This command will tell you what directory you are currently in
  - **ls [directory] [flags]** = "List": This command will list all files, folders, applications, etc. in your chosen **directory**, you can also add flags such as -l, -h, etc. If the **directory** parameter is left blank (for example if you were to just enter "ls"), then all files, etc. in your current (working) directory will be listed.
  - **cd [directory]** = "Change directory": This command will change your working directory to whatever you enter in **directory**. If the directory parameter is left blank, your working directory will be changed to your home directory.
  - **cd ..** = Your working directory will be changed to the folder directly above your current directory. For example, if you are in _~/.../folderA/folderB_, executing the command \"**cd ..**\" will change your working directory to ~/.../folderA
  - **cd \~** = Your working directory will be changed to your home directory.
* Navigate to the folder where you would like to copy the contents of this repository into.
* Clone the repository
  - Click the big green button above that says **\"<> Code\"**
  - Click **Local**
  - Click **HTTPS**
  - Copy the URL
  - In Git Bash, run the command **git clone \[Copied URL\]**

### Making changes to and and updating files in this repository
* In the _Microsoft Visual Basic for Applications_ window, in your project, open the \"Modules\" folder
* Right click on \[Insert module name here\]
* Click \"Export File...\"
* Navigate to where you have saved this repo saved in your file system
* Save the file
* You will get a warning saying that the file already exists, and asking if you want to replace it with this new version. Click Yes
* Export all Userforms using the same method as above.
* In Git Bash, run the command **\"git commit \*\"**
* A new vim window will pop up in Git Bash. Write your update log in here. Once you are done, save and close the file by typing **\"\:wq\"**
* In Git Bash, run the command **\"git push\"**













   
