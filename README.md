## Gmail-to-xlsx
This project reads your N number of mails and creates a xlsx file under sender information.

# Packages to be installed

`pip install imaplib`

`pip install python-dateutil`

`pip install openpyxl`

`pip install email`


# Pre-configuration to be done
Open your gmail account -> Settings -> Security -> find "Less secure app access" and Turnn it ON.(Don't worry this is just to enable that your script can have  access for fetching details).


# How to Run
Enter the gmail/G-suite id and password in line 7 and 8 respectively.

Enter the number of mails to be readed from your account in line 21 (Ex. N=10)

Now run the program after a while, a file named "Converted.xlsx" is created in which it has the details of sender name, email id, sent date, subject.


