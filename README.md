**This is a program that will allow to convert the Tally Transaction Log Reports to a much readable format**
This program also helps to identify possible errors in the report, as it iterates the transaction log report looking for common indicators of discrepancies or problems.

![image](https://github.com/Mrgoodkato/transaction_log/assets/81311715/2bc4a2b9-cd77-42ed-91e6-0835ebb107f4)


**INSTRUCTIONS**
1- When first using this program, create the following folder in the main folder (output and temp):

![image](https://github.com/Mrgoodkato/transaction_log/assets/81311715/1a63fda0-4d79-460a-a2f5-d63d26e3fd49)

2- Move the transaction log reports in **.xlsx** format to the files folder (very important they are in this format) without any changes done to the files. This program is designed to work with the naturally created .xlsx files from ACLEYNK.

![image](https://github.com/Mrgoodkato/transaction_log/assets/81311715/077e666a-8374-4999-920c-88692b54c787)

You will find a default file in the files folder for testing called banana_transactionlog.xlsx, you can remove it if you want.

3- Use your terminal or powershell or command line to run the python script using python app.py.
(You don't need to specify the name of the files, the program takes all files under the files folder and converts them in different interpretation files and folders in the output folder)

![image](https://github.com/Mrgoodkato/transaction_log/assets/81311715/13dff378-378e-4363-acd2-a613de6661df)

4- Check the output folder for the main interpretation files and the folder for a more differentiated analysis by tally out.

![image](https://github.com/Mrgoodkato/transaction_log/assets/81311715/1c2a8c69-10b6-4c73-9e0e-a030dc49c3f1)
