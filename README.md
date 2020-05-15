# Resume-CoverLetter-Filler
Automatically Fills Resume and Cover Letter with information

*IMPORTANT*: Only compatable with .docx files, not .doc

When this Python program is run, you are prompted to choose templates for your resume and cover letter.
The templates can have the following text in them:

[DATE] - The current date
[HIRINGMANGER] - The name of the hiring manager
[ADDRESS1] - The first address line of the company
[ADDRESS2] - The second address line of the company
[POSITION] - The position you're applying for
[COMPANY] - The name of the company
[GPA] - Your current GPA.

Those key terms will be replaced with the information filled in the program.
For example, the following text will be changed as shown with the inputted Position and Company:
  "...seeking a position as [POSITION] with [COMPANY]."
  "...seeking a position as a Software Engineer Intern with Software Company."
  
The program creates a new folder named after the company. This folder is placed in the same folder as the template documents.
The modified documents and their PDF versions are placed in the new folder. If the files already exist, copies are placed in the new
folder as well.
