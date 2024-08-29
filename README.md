# Adoption Score Generator
This module scrapes adoption scores from a list of clients and pastes them into another sheet.

## Set-up
To use this software, ensure that
1. Install `python` and `git` from the internet

2. In your terminal, install pip with 
   ```bash
   python -m ensurepip --upgrade
3. Clone the repository
   ```bash
   git clone https://github.com/wheesinsty/adoption-scores.git
4. Go to the repository
   ```bash
   cd [insert path to the repository]
5. Install dependencies
   ```bash
   pip install -r requirements.txt

## Usage
1. In the terminal, enter
   ```bash
   python adoption_scores.py
2. Enter the path to the excel sheet.

## Program not working?
Potential issues and how to resolve them:

1. Permission error
   How to resolve: Ensure that the excel sheet is closed.

2. Sheet not found
   How to resolve: Ensure that the worksheets are called "Adoption score" and "Report status"

## Contribution
Thank you to Herman and Clarice for giving me the opportunity to complete this project and contribute to the team :)
