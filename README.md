# GSTR-Download-Automation
Hey all, 
For starters, I am not a coder, I am an Advocate by profession and do this as a hobby (obviously with tremendous levels of help from YouTube videos and GPTs). 
While working at an audit firm I found one particular task of downloading multiple files very redundant and time-consuming. 
To tackle this problem I turned to coding and found a solution using selenium and allied py libraries.

How does this work:

For this to work you will have to create three Excel files (A message to coders, please forgive me for any mistakes I might have committed in this code, I am open to learning more from all of you)
Step 1: It takes all the information from the Excel provided via a path (I have marked those places where an Excel path is to be provided along with the format of the excel)

Step 2: It uses selenium to web scrape the website (path to which is already provided and need not be changed)

Step 3: It then feeds the relevant information into the necessary places using various classes

Step 4: It then loops the code till information from all the 12 months in a Financial Year is scrapped (Please note that the Financial Year should be as given on the website, for example, if the financial year in the GST website is "2022 - 23" then the same format must be used in the Excel and likewise it must be followed for the period and quarters as well, else the code will not work. Also, kindly note that for the code to loop, you will have to provide the information relating to financial year, quarter, and period once in the login and password Excel and 11 times in the second Excel (I still haven't figured out how to use user input :P)

Step 5: Eradicate redundant requesting and downloading tasks
