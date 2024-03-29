import re
import os
import pandas as pd
import numpy as np
import util
import openpyxl as op



def entryParseXL(entries):
    """Function that works by searching the specified list of file paths, and parsing only the ones with xlsx extension via Regex

    Args:
        entries (list): List of paths in the folder to be used to gather all the xl files for transaction log

    Returns:
        list: List of xlsx file names to be used in the creation of the database
    """
    result = []
    
    for entry in entries:
        regx = re.search('\.xlsx$', entry)
        if(regx):
            result.append(entry)
            
    return result

def databaseCreation(entries):
    """Function that re-interprets the original files input and returns a dictionary object which each entry is one of the interpreted reports
Indexed by the filename of that report

    Args:
        entries (list): List of filenames in xlsx for all the transactions log reports in the files folder

    Returns:
        dict: Dictionary object containing each of the dataframes from the transaction log reports gathered from the files folder (already interpreted)
    """
    database_logs = {}

    index = 1
    for entry in entries:

        xl_name = 'TallyTransactionLogReport'
        util.setupExcel(f'files/{entry}')
        
        try:
            log_df = pd.read_excel('temp/modified.xlsx', sheet_name=xl_name)
        except:
            print('Error reading the excel file, please confirm the file is a valid transaction log file from ACELYNK')
            return 0

        if(os.path.exists('temp/modified.xlsx')):
            os.remove('temp/modified.xlsx')

        try:
            list_of_items = log_df.loc[:, 'Item Code'].unique()
            list_of_tallyouts = log_df.loc[:, 'DONumber\n'].unique()
            importer = log_df.loc[:, 'Importer Account'].unique()
        except:
            print('File is not a valid transaction log file, please review')
            return 0

        file_name = ''

        if(len(list_of_items) == 1 and len(list_of_tallyouts) > 1):
            file_name = 'Item No {}'.format(list_of_items[0])
        elif(len(list_of_items) > 1 and len(list_of_tallyouts) == 1):
            file_name = 'Tally Out {}'.format(list_of_tallyouts[0])
        elif(len(importer) == 1 and len(list_of_items) > 1 & len(list_of_tallyouts) > 1):
            file_name = 'Importer {}'.format(importer[0])
        else:
            file_name = 'Various'

        interpreted_df = pd.DataFrame(columns=[
            'Action', 'Item No', 'Tally In', 'Tally Out', 'Qty Tally In', 'Qty removed', 'Qty rolled-back', 'New Total Tally In', 'Date'
        ])

        for row in range(len(log_df)):

            action = 'Tally out movement'
            rollback_qty = log_df.loc[row, 'Rolled Back Updated Qty']
            tally_in_result = log_df.loc[row, 'Tally In Updated Qty']
            action_qty = log_df.loc[row, 'Tally In Deducted Qty']
            tally_in_no = log_df.loc[row, 'Tally In']
            tally_out_no = log_df.loc[row, 'DONumber\n']
            item_no = log_df.loc[row, 'Item Code']
            date = log_df.loc[row, 'Created Date']

            if(np.isnan(rollback_qty)):
                tally_in_qty = tally_in_result + action_qty
                total = tally_in_qty - action_qty
                listofchanges = [action, item_no, tally_in_no, tally_out_no, tally_in_qty, action_qty, 0, total, date]
            else:
                action = 'Rollback movement'
                tally_in_qty = rollback_qty - action_qty
                total = rollback_qty
                listofchanges = [action, item_no, tally_in_no, tally_out_no, tally_in_qty, 0, action_qty, total, date]

            interpreted_df.loc[row] = listofchanges

        database_logs['DataFrame {} - {}'.format(index, file_name)] = interpreted_df
        index += 1
        
    return database_logs

def processDatabase(df):
    """This function works with the dataframe of each tally out to pull out each action done and the information related to each action (rollback, tally out)

    Args:
        df (pd.dataframe): Dataframe pertaining to a tally transaction log report interpreted

    Returns:
        pd.dataframe: Final log interpretation of all the actions taken in the original transaction log report
    """

    storage_object = {}
    storage_dataframe = pd.DataFrame(columns=['Action', 'Indexes', 'Item Nos', 'Tally Ins', 'Qty', 'Error', 'Err-location'])

    error_check_list = []
    key_marker = []
    key_indx = 0

    index = 0

    for key, tallyout in df.iterrows():

        #Check for the quantity to be stored
        qty = 0

        if(tallyout.iloc[5] == 0):
            qty = tallyout.iloc[6]
        else:
            qty = tallyout.iloc[5]

        #Creation of the storage_object in first iteration
        if(len(key_marker) == 0):
            
            storage_object[index] = util.modifyStorageObject(
                {
                    'tallyout': tallyout,
                    'key': key,
                    'qty': qty,
                    'err': False,
                    'err-location': []

                }
            )

            error_check_list.append(str(tallyout.iloc[1]) + str(tallyout.iloc[2]))

        #Check if the previous action was the same as the current one, if different execute
        elif(df.loc[key_marker[key_indx-1], 'Action'] != tallyout.iloc[0]):

            error_check_list = []

            index += 1
            storage_object[index] = util.modifyStorageObject(
                {
                    'tallyout': tallyout,
                    'key': key,
                    'qty': qty,
                    'err': False,
                    'err-location': []
                }
            )

            error_check_list.append(str(tallyout.iloc[1]) + str(tallyout.iloc[2]))

        else:
            
            storage_object[index]['Indexes'].append(key)
            storage_object[index]['Item Nos'].append(tallyout.iloc[1])
            storage_object[index]['Tally Ins'].append(tallyout.iloc[2])

            #Get rid of duplicates
            storage_object[index]['Item Nos'] = util.uniqueList(storage_object[index]['Item Nos'])
            storage_object[index]['Tally Ins'] = util.uniqueList(storage_object[index]['Tally Ins'])
            
            storage_object[index]['Qty'] += qty
            
            error_check_list.append(str(tallyout.iloc[1]) + str(tallyout.iloc[2]))

            if(util.duplicateFinder(error_check_list)):
                storage_object[index]['Error'] = True
                storage_object[index]['Err-location'].append(str(tallyout.iloc[1]) + ' ' + str(tallyout.iloc[2])) 


        key_marker.append(key)
        key_indx += 1

        storage_dataframe.loc[index] = storage_object[index]

    return storage_dataframe


def saveFinalAnalysis(df, file_name, tallyout):
    """Saves all tally out dataframes into separate files over separate folders according to the db object

    Args:
        df (pd.dataframe): Dataframe object to be saved in excel format
        file_name (str): File name
        tallyout (str): Tally out number
    """

    if(not os.path.isdir('output/{}'.format(file_name))):
        os.makedirs('output/{}'.format(file_name))
        
    file_path = 'output/{}/{}_final_analysis.xlsx'.format(file_name, tallyout)
    df.to_excel(file_path, index=False)

    workbook = op.load_workbook(file_path)
    sheet = workbook.active
    sheet.column_dimensions['A'].width = 20
    sheet.column_dimensions['B'].width = 20
    sheet.column_dimensions['C'].width = 20
    sheet.column_dimensions['D'].width = 20
    sheet.column_dimensions['E'].width = 20
    sheet.column_dimensions['F'].width = 20
    sheet.column_dimensions['G'].width = 20
    
    workbook.save(file_path)
    print(f'Saved file to - {file_path}')

def finalInterpretationSave(db):
    """This function takes the full db dictionary object containing all the dataframes gathered and saves them in an individual log interpretation file in the output folder

    Args:
        db (dict): Dictionary containing all log interpretation dataframes
    """
    database = {}
    
    for entry in db:

        db_entry = db[entry]
        db_entry['Errors'] = ''

        tally_out_list = db_entry.loc[:, 'Tally Out'].unique()

        tally_out_list

        for tout in tally_out_list:

            database[tout] = processDatabase(db_entry.loc[db_entry['Tally Out'] == tout])

            for error in database[tout].loc[:, ['Indexes', 'Error']].iterrows():

                if error[1]['Error'] == True:

                    db_entry.loc[error[1]['Indexes'], 'Error'] = True

            saveFinalAnalysis(database[tout], entry, tout)
        
        file_path = 'output/{}_log_interpretaion.xlsx'.format(entry)
        db_entry.to_excel(file_path, index=False)

        workbook = op.load_workbook(file_path)
        sheet = workbook.active
        sheet.column_dimensions['A'].width = 20
        sheet.column_dimensions['B'].width = 20
        sheet.column_dimensions['C'].width = 20
        sheet.column_dimensions['D'].width = 20
        sheet.column_dimensions['E'].width = 20
        sheet.column_dimensions['F'].width = 20
        sheet.column_dimensions['G'].width = 20
        sheet.column_dimensions['H'].width = 20
        sheet.column_dimensions['I'].width = 25
        
        util.excelStyler(sheet)

        workbook.save(file_path)

        print(f'Saved interpretation to - {file_path}')