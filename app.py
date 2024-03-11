import os
import df_processing as dfp

entries = dfp.entryParseXL(os.listdir('files'))
db = dfp.databaseCreation(entries)
if(db == 0): print('Fatal error, please restart the program')
else: 
    dfp.finalInterpretationSave(db)