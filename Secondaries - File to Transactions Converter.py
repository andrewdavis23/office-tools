from tkinter import *
from tkinter import filedialog
from datetime import timedelta
import pandas as pd
import os

prog_size = 0
program_folder = ''
program_file = ''

df_event = pd.DataFrame()
df_store = pd.DataFrame()
df_item = pd.DataFrame()

def standard_sec_logic(row):
    if row['CasePack'] == 1:
        return row['CasePack'] * 3
    elif row['CasePack'] * 2 > row['Capacity']:
        return row['CasePack']
    else:
        return row['CasePack'] * 2

def sec_logic_1(row):
    if row['CasePack'] == 1:
        return 60
    else:
        return row['CasePack'] * 5

def print_message(message, delete):
    # this function prints the message variable to tkinter, delete=True replaces old text
    if delete:
        files_txt.config(state=NORMAL)
        files_txt.delete('1.0', END)
        files_txt.insert(END, message)
        files_txt.config(state=DISABLED)
    else:
        files_txt.config(state=NORMAL)
        files_txt.insert(END, message)
        files_txt.config(state=DISABLED)

def load_program():
    global prog_size
    message = ''
    global program_folder
    global program_file
    global df_event
    global df_store
    global df_item
    global sec_logic

    # User picks program file, return if user cancels fialdialog ''
    program_file = filedialog.askopenfilenames()
    if program_file == '': return
    program_file = program_file[0]

    # extract folder and program name
    program_folder = os.path.dirname(program_file)

    # Process each sheet in workbook
    try:
        try:
            df_event = pd.read_excel(program_file, sheet_name='tbl_Event', names=['Start_Date','End_Date','Event','Half'])
        except:
            df_event = pd.read_excel(program_file, sheet_name=0, names=['Start_Date','End_Date','Event','Half'])

        try:
            df_store = pd.read_excel(program_file, sheet_name='tbl_Master_Filter', dtype={'BA_Filter':str}, names=['Div','StoreNum','BA_Filter'])
        except:
            df_store = pd.read_excel(program_file, sheet_name=1, names=['Div','StoreNum','BA_Filter'], dtype={'BA_Filter':str})
        
        try:
            df_item = pd.read_excel(program_file, sheet_name='tbl_Pog_Capacity', dtype={'Size':str}, names=['Half','Size','Location','Division','Description','ItemNum','Capacity','CasePack','Quantity'])
        except:
            df_item = pd.read_excel(program_file, sheet_name=2, names=['Half','Size','Location','Division','Description','ItemNum','Capacity','CasePack','Quantity'], dtype={'Size':str})
            
        message += f'Program loaded: {df_event.iloc[0,2]}'
        message += '\nTable (rows, columns) [column names]:\n'
        message += '+ Event {} {}\n'.format(df_event.shape,list(df_event.columns.values))
        message += '+ Store {} {}\n'.format(df_store.shape,list(df_store.columns.values))
        message += '+ Item  {} {}\n'.format(df_item.shape,list(df_item.columns.values))

        # CLEAN 1: Check kraft filter for not null, then drop column
        try:
            if not(df_store['Kraft_Filter'].isna().all()):
                message += '\nSTOP! Kraft_Filter is not null.\nDo this program manually.'
                print_message(message, False)
            df_store = df_store.drop(['Kraft_Filter'], axis=1)
        except:
            None

        # CLEAN 2: In the case where size A = 'Yes' and size B = 'YES'
        df_item['Size'] = df_item['Size'].str.upper()
        df_store['BA_Filter'] = df_store['BA_Filter'].str.upper()

        # CLEAN 3: Remove stores that are not part of the program
        df_store = df_store[df_store['BA_Filter'] != 'NO']
        df_store = df_store.dropna()

        # PREPROCESS: Calculate secondary capacity from capacity, casepack then drop columns
        if sec_logic.get() == 1:
            df_item['Sec_Cap'] = df_item.apply(standard_sec_logic, axis=1)
            df_item = df_item.drop(['Capacity','CasePack'], axis=1)
            message += '\nStandard Secondary Logic Used'
        elif sec_logic.get() == 2:
            df_item['Sec_Cap'] = df_item.apply(sec_logic_1, axis=1)
            df_item = df_item.drop(['Capacity','CasePack'], axis=1)
            message += '\nAlt Secondary Logic #1 Used'
        else:
            message += '\nChoice of secondary logic ({}) is not coded into program.'.format(sec_logic.get())

        # Show the user data types and category values prior to merging the tables on store-size category
        message += '\n+ Store Size Categories: {} Data type: {}'.format(df_store['BA_Filter'].unique(),df_store.dtypes['BA_Filter'])
        message += '\n+ Item Size Categories : {} Data type: {}\n'.format(df_item['Size'].unique(),df_item.dtypes['Size'])

    except BaseException as em:
        message += '\nLoad Fail\n\tException Message: {}\n'.format(em)

    # some stats about the program
    prog_size = len(list(df_event.Half.unique()))
    message += '\n1 part or 2 part program: {}'.format(prog_size)

    print_message(message, True)

def produce_file():
    # clean data, calculate secondary capacity, create trans files for program, save files
    global program_folder
    global prog_size
    global df_event
    global df_store
    global df_item
    message = ''
    trans_cols = ['STR_MAINT_TRANS_ID','STR_ID','STR_NUM','REF_NUM','STR_MAINT_TRANS_CD','STR_MAINT_NUM','STR_MAINT_CD','STR_MAINT_CHAR_TXT','PRCS_CD','USER_ID','CRT_TS','PRCS_TS','Program_Name']
    
    # extract event info: name, start, end
    event_name = df_event.iloc[0,2]

    if prog_size == 1:
        # Inner join the stores and items on division and size where half is first/second/both and secondary capacity > 0.
        df_trans_in = pd.merge(df_item.loc[(df_item['Sec_Cap']>0) & (df_item['Half']==1)], df_store, how='inner', left_on=['Division','Size'], right_on=['Div','BA_Filter'])
        
        # drop columns not needed in transmission file
        df_trans_in = df_trans_in.drop(['Half','Size','Division','Div','BA_Filter'], axis=1)

        # add/rename columns needed in transmission file
        df_trans_in = df_trans_in.rename(columns={'StoreNum': 'STR_NUM', 'ItemNum': 'REF_NUM', 'Sec_Cap':'STR_MAINT_NUM'})
        df_trans_in.insert(0,'STR_MAINT_TRANS_ID','')
        df_trans_in['STR_ID'] = df_trans_in['STR_NUM']
        df_trans_in['STR_MAINT_TRANS_CD'] = 'IPSC'
        df_trans_in['STR_MAINT_CD'] = 'N'
        df_trans_in['STR_MAINT_CHAR_TXT'] = '(null)'
        df_trans_in['PRCS_CD'] = 'T'
        df_trans_in['USER_ID'] = 'nuajd15'
        df_trans_in['CRT_TS'] = '(null)'
        df_trans_in['PRCS_TS'] = '(null)'
        df_trans_in['Program_Name'] = event_name

        # reorder columns
        df_trans_in = df_trans_in.reindex(columns = trans_cols)

        # copy out file from in
        df_trans_out = df_trans_in.copy()
        df_trans_out['STR_MAINT_NUM'] = 0
        df_trans_out['PRCS_TS'] = '(null)'
        df_trans_out['Program_Name'] = event_name

        # create schedule for whiteboard notes
        p_in = df_event.loc[0,'Start_Date']
        p_in = p_in - timedelta(days=1)
        p_in = p_in.strftime('%m/%d/%Y')

        p_out = df_event.loc[0,'End_Date']
        p_out = p_out - timedelta(weeks=2) - timedelta(days=1)
        p_out = p_out.strftime('%m/%d/%Y')

        with open(os.path.dirname(program_folder)+'/schedule.txt', 'a') as file:
            file.write('{}\t{}\t{}\n'.format(event_name, p_in, p_out))

        # save files
        fp = os.path.dirname(program_folder) + '/in and out files/' + df_event.iloc[0,2]
        df_trans_in.to_csv('{} IN.txt'.format(fp), sep='\t', index=False)
        df_trans_out.to_csv('{} OUT.txt'.format(fp), sep='\t', index=False)
        message += '\nSTART {}\t\t\tEND {}'.format(p_in, p_out)    
        message += '\n\nIN-file saved to\n   {} IN.txt\n   # of rows = {}'.format(fp,df_trans_in.shape[0])
        message += '\n\nOUT-file saved to\n   {} OUT.txt\n   # of rows = {}'.format(fp,df_trans_out.shape[0])

    elif prog_size == 2:
        # HalfWord is first, second, or both.  Merge to add HalfWord to item dataframe
        df_halves = pd.read_excel(program_file, sheet_name='Items by Half', usecols='A:B')
        df_halves.columns = ['ItemNum', 'HalfWord']
        df_item = pd.merge(df_item, df_halves, how='inner', left_on='ItemNum', right_on='ItemNum')

        # Inner join the stores and items on division and size where half is first/second/both and secondary capacity > 0.
        # first and both
        df_trans_first_in = pd.merge(df_item.loc[(df_item['Sec_Cap']>0) & ((df_item['HalfWord']=='First')|(df_item['HalfWord']=='Both'))], df_store, how='inner', left_on=['Division','Size'], right_on=['Div','BA_Filter'])
        # first
        df_trans_first_out = pd.merge(df_item.loc[(df_item['Sec_Cap']>0) & (df_item['HalfWord']=='First')], df_store, how='inner', left_on=['Division','Size'], right_on=['Div','BA_Filter'])
        # second
        df_trans_second_in = pd.merge(df_item.loc[(df_item['Sec_Cap']>0) & (df_item['HalfWord']=='Second')], df_store, how='inner', left_on=['Division','Size'], right_on=['Div','BA_Filter'])
        # second and both
        df_trans_second_out = pd.merge(df_item.loc[(df_item['Sec_Cap']>0) & ((df_item['HalfWord']=='Second')|(df_item['HalfWord']=='Both'))], df_store, how='inner', left_on=['Division','Size'], right_on=['Div','BA_Filter'])

        # list of dataframes
        trans_list = [df_trans_first_in, df_trans_first_out, df_trans_second_in, df_trans_second_out]

        # makes the columns null, leave blank instead?
        df_trans_first_in['PRCS_TS'] = '(null)'
        df_trans_first_out['PRCS_TS'] = '(null)'
        df_trans_second_in['PRCS_TS'] = '(null)'
        df_trans_second_out['PRCS_TS'] = '(null)'
        
        # change secondary value to 0 in removal files
        df_trans_first_out.drop('Sec_Cap', axis=1, inplace=True)
        df_trans_first_out['STR_MAINT_NUM'] = 0
        df_trans_second_out.drop('Sec_Cap', axis=1, inplace=True)
        df_trans_second_out['STR_MAINT_NUM'] = 0

        # columns that each df should have
        for i in trans_list:
            i.drop(['Half','Size','Division','Div','BA_Filter','HalfWord','Location','Quantity','Description'], axis=1, inplace=True)
            i.rename(columns={'StoreNum': 'STR_NUM', 'ItemNum': 'REF_NUM', 'Sec_Cap':'STR_MAINT_NUM'}, inplace=True)
            i.insert(0,'STR_MAINT_TRANS_ID','',allow_duplicates=True)
            i.insert(0,'STR_MAINT_CD','N',allow_duplicates=True)
            i.insert(0,'STR_MAINT_CHAR_TXT','(null)',allow_duplicates=True)
            i.insert(0,'PRCS_CD','T',allow_duplicates=True)
            i.insert(0,'USER_ID','nuajd15',allow_duplicates=True)
            i.insert(0,'CRT_TS','(null)',allow_duplicates=True)
            i.insert(0,'STR_ID', i['STR_NUM'],allow_duplicates=True)
            i.insert(0,'STR_MAINT_TRANS_CD','IPSC',allow_duplicates=True)
            i.insert(0,'Program_Name',event_name,allow_duplicates=True)

        df_trans_first_in = df_trans_first_in.reindex(columns = trans_cols)
        df_trans_first_out = df_trans_first_out.reindex(columns = trans_cols)
        df_trans_second_in = df_trans_second_in.reindex(columns = trans_cols)
        df_trans_second_out = df_trans_second_out.reindex(columns = trans_cols)

        # create schedule for whiteboard notes
        p_first_in = df_event.loc[0,'Start_Date']
        p_first_in = p_first_in - timedelta(days=1)
        p_first_in = p_first_in.strftime('%m/%d/%Y')

        p_first_out = df_event.loc[0,'End_Date']
        p_first_out = p_first_out - timedelta(days=14)
        p_first_out = p_first_out.strftime('%m/%d/%Y')

        p_second_in = df_event.loc[1,'Start_Date']
        p_second_in = p_second_in - timedelta(days=1)
        p_second_in = p_second_in.strftime('%m/%d/%Y')

        p_second_out = df_event.loc[1,'End_Date']
        p_second_out = p_second_out - timedelta(days=14)
        p_second_out = p_second_out.strftime('%m/%d/%Y')

        with open(os.path.dirname(program_folder) + 'schedule.txt', 'a') as file:
            file.write('{} 1\t{}\t{}\n'.format(event_name, p_first_in, p_first_out))
            file.write('{} 2\t{}\t{}\n'.format(event_name, p_second_in, p_second_out))
      
        # save files
        fp = os.path.dirname(program_folder) + '/in and out files/' + df_event.iloc[0,2]
        df_trans_first_in.to_csv('{} FIRST IN.txt'.format(fp), sep='\t', index=False)
        df_trans_first_out.to_csv('{} FIRST OUT.txt'.format(fp), sep='\t', index=False)
        df_trans_second_in.to_csv('{} SECOND IN.txt'.format(fp), sep='\t', index=False)
        df_trans_second_out.to_csv('{} SECOND OUT.txt'.format(fp), sep='\t', index=False)
        message += '\nFIRST HALF\n'
        message += '\tSTART\t{}\t\t\tEND\t{}\n'.format(p_first_in, p_first_out)    
        message += '\nSECOND HALF\n'
        message += '\tSTART\t{}\t\t\tEND\t{}\n'.format(p_second_in, p_second_out)    
        message += '\tIN\t{}\t\t\tOUT\t{}'.format(p_second_in, p_second_out)
        message += '\n\nFILES SAVED TO {}'.format(fp)
        message += '\n\tFIRST IN\t{} rows\n\tFIRST OUT\t{}\n\tSECOND IN\t{}\n\tSECOND OUT\t{}'.format(df_trans_first_in.shape[0],df_trans_first_out.shape[0],df_trans_second_in.shape[0],df_trans_second_out.shape[0])

    else:
        message += '\nProgram size not 1 or 2. See the tbl_Event tab in the progam file.\n'

    print_message(message, False)

root = Tk()
root.title('Secondary Transaction Creator')
root.geometry('900x450')
sec_logic = IntVar()

left_frame = Frame(root)
left_frame.pack(side=LEFT)

right_frame = Frame(root)
right_frame.pack(side=LEFT)

right_up = Frame(right_frame)
right_up.pack(side=TOP)

right_down = Frame(right_frame)
right_down.pack(side=BOTTOM)

# LEFT FRAME
    
R1 = Radiobutton(left_frame, text = 'Standard Logic', variable = sec_logic,
        value = 1, background = "light blue")
R1.grid(row=0, column=0)

R2 = Radiobutton(left_frame, text = 'Alternate Logic', variable = sec_logic,
        value = 2, background = "light blue")
R2.grid(row=1, column=0)

L1 = 'Standard Logic:\n  PK = 1 ⇒ SEC = PK*3\n  PK*2 > CAP ⇒ SEC = PK\n  Else SEC = PK*2'
L2 = 'Alt Logic #1:\n  PK=1 ⇒ SEC=60\n  Else SEC=PK*5'

L1 = Label(left_frame, text=L1)
L2 = Label(left_frame, text=L2)

L1.grid(row=0, column=1)
L2.grid(row=1, column=1)

# RIGHT FRAME

files_txt = Text(right_up)
files_txt.pack()

load_button = Button(right_down, text='Select Secondary\nProgram File', command=load_program)
load_button.grid(row=0, column=0, sticky='nsew', padx=2)

pro_button = Button(right_down, text='Produce\nTransaction Files', command=produce_file)
pro_button.grid(row=0, column=1, sticky='nsew', padx=2)

root.mainloop()
