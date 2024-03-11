import openpyxl as op

#Function used to modify the storage object, provide initial values
def modifyStorageObject(data):

    return {
            
            'Action': data['tallyout'][0],
            'Indexes': [data['key']],
            'Item Nos': [data['tallyout'][1]],
            'Tally Ins': [data['tallyout'][2]],
            'Qty': data['qty'],
            'Error': data['err'],
            'Err-location': data['err-location']

        }

#Function that creates a unique value list from a list that might have duplicate values
def uniqueList(list):

    unique_list = []

    for itm in list:
        if itm not in unique_list:
            unique_list.append(itm)
    
    return unique_list

#Function returns True if there are duplicates in the list, False if there are no duplicates
def duplicateFinder(list):

    unique_list = []

    for itm in list:
        if itm not in unique_list:
            unique_list.append(itm)
        else:
            return True
        
    return False

#Function that removes the necessary rows and columns from the Transaction log report
def setupExcel(path):

    wb = op.load_workbook(path)
    ws = wb.active

    ws.unmerge_cells('B2:D2')
    ws.unmerge_cells('B4:J4')
    
    num = 6
    while True:
        try:
            ws.unmerge_cells(f'D{num}:E{num}')
            ws.unmerge_cells(f'J{num}:K{num}')
            num += 1
        except:
            break

    ws.delete_rows(1, 5)
    ws.delete_cols(1, 1)
    ws.delete_cols(4, 1)
    ws.delete_cols(9, 1)

    wb.save('temp/modified.xlsx')