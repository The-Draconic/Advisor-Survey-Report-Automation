# Advisor-Survey-Report-Automation
This project takes the given file, reads it and takes its data to put into three different kinds of file:
1. **3 Year Combined Averages**, the 3 Year rolling data sheet of all the surveys for all the advisors.
2. **All Adviser 3 Year Stats**, the summary of the data in the form of averages for all the advisors.
3. **Individual Advisor Sheets**, an individual sheet for each advisor to see all of their survey responses, the averages and total responses number.

## Instructions for Use
### Acquiring and Setting-up the Code
If you've been sent a zip file with the code and files, then you've already accomplished acquiring the code. But now you need to place this in a folder and run this code from Visual Studio Code. 

If you're reading this from GitHub and haven't gotten the code yet, open your terminal (command prompt or Git Bash), navigate to the directory where you want to store the repository, and run the following command: git clone <repository_url>. You can get the url from the GitHub page, click on **Code** and either grab the HTTP or SSH url. Though it may be simpler to use the HTTP url.

When you get the code, make sure to-in order to see all the files-you open the folder which you placed it all with VSCode, or otherwise navigate to this folder. Instructions to run the code in the next section below.

### Running the Code
When you have VSCode open and you see the code and files, follow these instructions:
1. Make sure that, besides the code files, you have the **3 Year Combined Averages** file (from the past year) and the Qualtrics file in the folder on the same level as main.py. If you got the code from GitHub, you'll need to make sure you do that.
2. Right-click on the folder with all the code and select '*Open in Integrated Terminal*'.
3. In the terminal, enter: python main.py
4. After entering that, you'll be prompted for the new year, enter the year of the survey data. (ex. 2024)
5. Next, you'll be prompted for what file you want to read from, copy and paste the name of the survey file, INCLUDE THE FILE EXTENSION. (ex. 2024 Advisor Survey_March 27, 2024_17.50.csv)
6. The code will then run the rest on its own, editing the **3 Year Combined Averages** file, create the **All Adviser 3 Year Stats** file and individual files for each advisor inside the folder called **Individual Advisor Data**.

## Notes for Future Programmers
If you cloned the GitHub repository to your local repository and are planning to push updated code to GitHub, DO NOT, and I repeat, DO NOT PUSH THE CODE with the files containing sensitive data. Take those files out of the repository folder and then push it.