# Adoption Score Generator
This module scrapes adoption scores from a list of clients and pastes them into another sheet.

## Set-up
To use this software, ensure that

1. Install python from the internet

2. In your terminal, install pip with 
`python -m ensurepip --upgrade`

3. Clone this repository with git clone 
`https://github.com/wheesinsty/adoption-scores.git`

4. Right click the Google Chrome app and select Properties. Then, set the Target: to 
`"C:\ProgramFiles\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222`

5. Check that the repository is in the Downloads folder, then navigate to this directory in your terminal by entering 
`cd [path]`

6. Install the necessary dependencies with 
`pip install -r requirements.txt`

## Run the program
1. Open terminal and navigate to the directory 
`cd [path]`

2. Run the file 
`python advisory.py`

3. Enter the path to the excel sheet

## Project doesn't work?
Potential issues and how to resolve them:

1. Permission error:
How to resolve: Ensure that the advisory template and the excel sheets are closed.

2. Sheet not found:
Ensure that the sheets are named "Adoption score" and "Report status"

## Credits
Thank you to my team for giving me this opportunity and for helping me along the way :`)`