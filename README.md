# Chime Chatroom Tool
This script automates the process of adding members to a chat room in Chime, without the using Chime API. This is intended for users who are unable to utilize certain methods from Chime API due to an absence of an enterprise account.

## Prerequisites
Chrome Version: 81


Python: 3.6^

## Setup
`pip install -r requirements.txt`


Replace the chrome driver in the driver folder with a version that's appropriate for your browser. The chromedriver version that is provided in this repository is for Chrome version 81.x 

## Configure

In config/parameters.ini:


1. Change the CHATROOM_URL parameter to the URL of your chat room. Ensure that you are either the owner or administrator of that chat room.


2. Change the CHROME_DRIVER_PATH parameter to your chromedriver's path if you are NOT using the default chromedriver that is provided here for Chrome version 81.x


3. Update the EMAIL_LIST_FILE parameter to the path of your excel file that contains the list of emails to be added to the chat room. Ensure that the emails are stored in individual cells under column 'A'. Refer to files/email_list.xlsx for the example.


4. Change the SHEET_NAME parameter to the name of your sheet that contains the list of emails. Default sheet name is 'Sheet1'


## Run

Ensure that you have configured properly in the previous section before running the script.


On your terminal, run the command:
`python chime-room-tool.py`


When chrome launches and navigates to Chime, login with your credentials. 


Once you have logged in, press Enter on your terminal. This will start a loop, where it adds every member from the excel sheet to the chat room. Once every member is added the script will stop and 2 new sheets will be created in the excel file: 'Failed' and 'Added'. The 'Failed' sheet contains the list of emails that were failed to be added to the chat room. The 'Added' sheet contains the list of names of the members that were added to the chat room.
