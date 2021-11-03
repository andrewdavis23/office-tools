from tkinter import *
# from numpy import right_shift
from tkcalendar import *
from tkinter import filedialog
from datetime import date
from datetime import timedelta
import pandas as pd
import os

prog_size = 0
program_folder = ''
program_file = ''
# CHANGE THIS MANUALLY - RADIO BUTTONS NOT WORKING
sec_logic = 1

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

    # User picks program file, return if user cancels fialdialog ''
    program_file = filedialog.askopenfilenames()
    if program_file == '': return
    program_file = program_file[0]

    # extract folder and program name
    program_folder = os.path.dirname(program_file)

    # Process each sheet in workbook
    try:
        df_event = pd.read_excel(program_file, sheet_name='tbl_Event')
        df_store = pd.read_excel(program_file, sheet_name='tbl_Master_Filter', dtype={'BA_Filter':str})
        df_item = pd.read_excel(program_file, sheet_name='tbl_Pog_Capacity', dtype={'Size':str})
        message += f'Program loaded: {df_event.iloc[0,2]}'
        message += '\n# of (rows, columns) loaded:\n'
        message += '   {} in Event Data\n'.format(df_event.shape)
        message += '   {} in Store Data\n'.format(df_store.shape)
        message += '   {} in Item Capacity Data\n'.format(df_item.shape)

        # CLEAN 1: Check kraft filter for not null, then drop column
        if not(df_store['Kraft_Filter'].isna().all()):
            message += '\nSTOP! Kraft_Filter is not null.\nDo this program manually.'
            print_message(message, False)
        df_store = df_store.drop(['Kraft_Filter'], axis=1)

        # CLEAN 2: In the case where size A = 'Yes' and size B = 'YES'
        df_item['Size'] = df_item['Size'].str.upper()
        df_store['BA_Filter'] = df_store['BA_Filter'].str.upper()

        # CLEAN 3: Remove stores that are not part of the program
        df_store = df_store[df_store['BA_Filter'] != 'NO']
        df_store = df_store.dropna()

        # PREPROCESS: Calculate secondary capacity from capacity, casepack then drop columns
        if sec_logic == 1:
            df_item['Sec_Cap'] = df_item.apply(standard_sec_logic, axis=1)
            df_item = df_item.drop(['Capacity','CasePack'], axis=1)
            message += '\n   Standard Secondary Logic Used'
        elif sec_logic == 2:
            df_item['Sec_Cap'] = df_item.apply(sec_logic_1, axis=1)
            df_item = df_item.drop(['Capacity','CasePack'], axis=1)
            message += '\n   Alt Secondary Logic #1 Used'
        else:
            message += '\n   Choice of secondary logic is not coded into program.'

        # Show the user data types and category values prior to merging the tables on store-size category
        message += '\n   Store Size Categories: {} Data type: {}'.format(df_store['BA_Filter'].unique(),df_store.dtypes['BA_Filter'])
        message += '\n   Item Size Categories : {} Data type: {}\n'.format(df_item['Size'].unique(),df_item.dtypes['Size'])

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

    if prog_size == 1:
        # join the stores and items where half
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
        #'2020/12/17:03:49:00 PM' pd.datetime being buggy
        # will have to get using look up half=2....
        p_start = df_event.loc[0,'Start_Date']
        p_start = p_start - timedelta(days=1)
        p_in = p_start.strftime('%Y/%m/%d')+':03:49:00 PM'
        df_trans_in['PRCS_TS'] = p_in

        # reorder columns
        df_trans_in = df_trans_in.reindex(columns = ['STR_MAINT_TRANS_ID','STR_ID','STR_NUM','REF_NUM','STR_MAINT_TRANS_CD','STR_MAINT_NUM','STR_MAINT_CD','STR_MAINT_CHAR_TXT','PRCS_CD','USER_ID','CRT_TS','PRCS_TS'])

        # copy out file from in
        df_trans_out = df_trans_in.copy()
        df_trans_out['STR_MAINT_NUM'] = 0
        p_end = df_event.loc[0,'End_Date']
        p_end = p_end - timedelta(weeks=2)
        p_out = p_end.strftime('%Y/%m/%d')+':03:49:00 PM'
        df_trans_out['PRCS_TS'] = p_out

        # save files
        fp = os.path.dirname(program_folder) + '/in and out files/' + df_event.iloc[0,2]
        df_trans_in.to_csv('{} IN.txt'.format(fp), sep='\t', index=False)
        df_trans_out.to_csv('{} OUT.txt'.format(fp), sep='\t', index=False)
        message += '\nSTART {}\t\t\tEND {}'.format(df_event.loc[0,'Start_Date'].strftime('%Y/%m/%d')+' '*12, df_event.loc[0,'End_Date'].strftime('%Y/%m/%d')+' '*12)    
        message += '\nIN    {}\t\t\tOUT {}'.format(p_in, p_out)    
        message += '\n\nIN-file saved to\n   {} IN.txt\n   # of rows = {}'.format(fp,df_trans_in.shape[0])
        message += '\n\nOUT-file saved to\n   {} OUT.txt\n   # of rows = {}'.format(fp,df_trans_out.shape[0])

    elif prog_size == 2:
        # HalfWord is first, second, both
        df_halves = pd.read_excel(program_file, sheet_name='Items by Half', usecols='A:B')
        df_halves.columns = ['ItemNum', 'HalfWord']
        df_item = pd.merge(df_item, df_halves, how='inner', left_on='ItemNum', right_on='ItemNum')

        # join the stores and items where half

        # first and both
        df_trans_first_in = pd.merge(df_item.loc[(df_item['Sec_Cap']>0) & ((df_item['HalfWord']=='First')|(df_item['HalfWord']=='Both'))], df_store, how='inner', left_on=['Division','Size'], right_on=['Div','BA_Filter'])
        # first
        df_trans_first_out = pd.merge(df_item.loc[(df_item['Sec_Cap']>0) & (df_item['HalfWord']=='First')], df_store, how='inner', left_on=['Division','Size'], right_on=['Div','BA_Filter'])
        # second
        df_trans_second_in = pd.merge(df_item.loc[(df_item['Sec_Cap']>0) & (df_item['HalfWord']=='Second')], df_store, how='inner', left_on=['Division','Size'], right_on=['Div','BA_Filter'])
        # second and both
        df_trans_second_out = pd.merge(df_item.loc[(df_item['Sec_Cap']>0) & ((df_item['HalfWord']=='Second')|(df_item['HalfWord']=='Both'))], df_store, how='inner', left_on=['Division','Size'], right_on=['Div','BA_Filter'])

        trans_list = [df_trans_first_in, df_trans_first_out, df_trans_second_in, df_trans_second_out]

        p_first_start = df_event.loc[0,'Start_Date']
        p_first_start = p_first_start - timedelta(days=1)
        p_first_start = p_first_start.strftime('%Y/%m/%d')+':03:49:00 PM'
        df_trans_first_in['PRCS_TS'] = p_first_start

        p_first_end = df_event.loc[0,'End_Date']
        p_first_end = p_first_end - timedelta(weeks=2)
        p_first_end = p_first_end.strftime('%Y/%m/%d')+':03:49:00 PM'
        df_trans_first_out['PRCS_TS'] = p_first_end

        p_second_start = df_event.loc[1,'Start_Date']
        p_second_start = p_second_start - timedelta(days=1)
        p_second_start = p_second_start.strftime('%Y/%m/%d')+':03:49:00 PM'
        df_trans_second_in['PRCS_TS'] = p_second_start

        p_second_end = df_event.loc[1,'End_Date']
        p_second_end = p_second_end - timedelta(weeks=2)
        p_second_end = p_second_end.strftime('%Y/%m/%d')+':03:49:00 PM'
        df_trans_second_out['PRCS_TS'] = p_second_end

        for i in trans_list:
            i.drop(['Half','Size','Division','Div','BA_Filter','HalfWord','Location'], axis=1, inplace=True)
            i.rename(columns={'StoreNum': 'STR_NUM', 'ItemNum': 'REF_NUM', 'Sec_Cap':'STR_MAINT_NUM'}, inplace=True)
            i.insert(0,'STR_MAINT_TRANS_ID','',allow_duplicates=True)
            i.insert(0,'STR_MAINT_CD','N',allow_duplicates=True)
            i.insert(0,'STR_MAINT_CHAR_TXT','(null)',allow_duplicates=True)
            i.insert(0,['PRCS_CD'],'T',allow_duplicates=True)
            i.insert(0,['USER_ID'],'nuajd15',allow_duplicates=True)
            i.insert(0,['CRT_TS'],'(null)',allow_duplicates=True)

        df_trans_first_in.reindex(columns = ['STR_MAINT_TRANS_ID','STR_ID','STR_NUM','REF_NUM','STR_MAINT_TRANS_CD','STR_MAINT_NUM','STR_MAINT_CD','STR_MAINT_CHAR_TXT','PRCS_CD','USER_ID','CRT_TS','PRCS_TS'])
        df_trans_first_out.reindex(columns = ['STR_MAINT_TRANS_ID','STR_ID','STR_NUM','REF_NUM','STR_MAINT_TRANS_CD','STR_MAINT_NUM','STR_MAINT_CD','STR_MAINT_CHAR_TXT','PRCS_CD','USER_ID','CRT_TS','PRCS_TS'])
        df_trans_second_in.reindex(columns = ['STR_MAINT_TRANS_ID','STR_ID','STR_NUM','REF_NUM','STR_MAINT_TRANS_CD','STR_MAINT_NUM','STR_MAINT_CD','STR_MAINT_CHAR_TXT','PRCS_CD','USER_ID','CRT_TS','PRCS_TS'])
        df_trans_second_out.reindex(columns = ['STR_MAINT_TRANS_ID','STR_ID','STR_NUM','REF_NUM','STR_MAINT_TRANS_CD','STR_MAINT_NUM','STR_MAINT_CD','STR_MAINT_CHAR_TXT','PRCS_CD','USER_ID','CRT_TS','PRCS_TS'])
             
    else:
        message += '\nProgram size not 1 or 2. Prep this one manually.\n'

    print_message(message, False)

root = Tk()
root.title('Secondary Transaction Creator')
root.geometry('900x450')

left_frame = Frame(root)
left_frame.grid(row=0, column=0)

right_frame = Frame(root)
right_frame.grid(row=0, column=1)

right_up = Frame(right_frame)
right_up.grid(row=0, column=0)

right_down = Frame(right_frame)
right_down.grid(row=1, column=0)

# LEFT FRAME
    
# R1 = Radiobutton(left_frame, text = 'Standard Logic', variable = sec_logic,
#         value = 1, indicator = 1,
#         background = "light blue")
# R1.grid(row=3, column=0)

# R2 = Radiobutton(left_frame, text = 'Alternate Logic', variable = sec_logic,
#         value = 2, indicator = 1,
#         background = "light blue")
# R2.grid(row=3, column=1)

L1 = 'Standard Logic:\n  PK=1 -> SEC=PK*3\n  PK*2>CAP -> SEC=PK\n  Else SEC=PK*2'
L2 = 'Alt Logic #1:\n  PK=1 -> SEC=60\n  Else SEC=PK*5'

L1 = Label(left_frame, text=L1)
L2 = Label(left_frame, text=L2)

L1.grid(row=4, column=0)
L2.grid(row=4, column=1)

# RIGHT FRAME

files_txt = Text(right_up)
files_txt.pack()

load_button = Button(right_down, text='Select Secondary\nProgram File', command=load_program)
load_button.grid(row=0, column=0, sticky='nsew', padx=2)

pro_button = Button(right_down, text='Produce\nTransaction Files', command=produce_file)
pro_button.grid(row=0, column=1, sticky='nsew', padx=2)

root.mainloop()
