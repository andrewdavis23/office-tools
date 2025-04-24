from tkinter import filedialog
import pandas as pd
import os

class SecProgClass:
    def __init__(self):
        self.df_event = pd.DataFrame()
        self.df_store = pd.DataFrame()
        self.df_item = pd.DataFrame()
        self.df_halves = pd.DataFrame()
        self.prog_size = 0
        self.file_name = ''
        self.folder = ''
        self.sec_logic = 1
        self.event_name = ''
        self.store_row_count = 0
        self.item_row_count = 0
        self.trans_cols = ['STR_ID','STR_NUM','REF_NUM','STR_MAINT_TRANS_CD','STR_MAINT_NUM','STR_MAINT_CD','STR_MAINT_CHAR_TXT','USER_ID','Program_Name']
        self.trans_files = []
        self.file_path = ''

def standard_sec_logic(row):
    if row['CasePack'] == 1:
        return row['CasePack'] * 3
    elif row['CasePack'] * 2 > row['Capacity']:
        return row['CasePack']
    else:
        return row['CasePack'] * 2

def load_event_data(program_file):
    col_names = ['Start_Date','End_Date','Event','Half']
    try:
        df = pd.read_excel(program_file, sheet_name='tbl_Event', names=col_names, header=0)
    except:
        print('Unexpected column names')
        df = pd.read_excel(program_file, sheet_name=0, names=col_names, header=0)
    return df

def load_store_data(program_file):
    col_names = ['LEGACY DIVISION', 'STORE', 'BA_Filter']
    try:
        df = pd.read_excel(program_file, sheet_name='tbl_Master_Filter', dtype={'BA_Filter':str})
    except:
        df = pd.read_excel(program_file, sheet_name=1, names=col_names, dtype={'BA_Filter':str}, header=0)
    return df

def load_item_data(program_file):
    col_names = ['Half','Size','Location','Division','Description','ItemNum','Capacity','CasePack','Quantity']
    try:
        df = pd.read_excel(program_file, sheet_name='tbl_Pog_Capacity', dtype={'Size':str})
    except:
        df = pd.read_excel(program_file, sheet_name=2, names=col_names, dtype={'Size':str}, header=0)
    return df

def load_half_data(program_file):
    col_names = ['ItemNum', 'HalfWord']
    try:
        df = pd.read_excel(program_file, sheet_name='Items by Half', usecols='A:B')
    except:
        df = pd.read_excel(program_file, sheet_name=3, names=col_names, header=0)
    return df

def load_program():
    SP = SecProgClass()

    # User picks program file, return if user cancels fialdialog ''
    fn = filedialog.askopenfilenames()
    if fn == '': 
        return None
    SP.file_name = fn[0]

    # extract folder and program name
    SP.folder = os.path.dirname(SP.file_name)

    # create data frames from sheets
    SP.df_event = load_event_data(SP.file_name)
    SP.df_store = load_store_data(SP.file_name)
    SP.df_item = load_item_data(SP.file_name)

    SP.event_name = SP.df_event.iloc[0,2]
    SP.prog_size = SP.df_event.iloc[0,3]
    SP.store_row_count = SP.df_store.shape[0]
    SP.item_row_count = SP.df_item.shape[0]

    # print data shape
    print('*' * 120)
    print('Program loaded: {}'.format(SP.event_name))
    print('+ Store {} Rows: {}'.format(list(SP.df_store.columns.values), SP.store_row_count))
    print('+ Item  {} Rows: {}'.format(list(SP.df_item.columns.values), SP.item_row_count))

    # Show the user data types and category values prior to merging the tables on store-size category
    print('+ Store Size Categories: {} Data type: {}'.format(SP.df_store['BA_Filter'].unique(), SP.df_store.dtypes['BA_Filter']))
    print('+ Item Size Categories : {} Data type: {}\n'.format(SP.df_item['Size'].unique(), SP.df_item.dtypes['Size']))

    return SP

def clean(SP):
    # CLEAN 1: Check kraft filter for not null, then drop column
    try:
        if not(SP.df_store['Kraft_Filter'].isna().all()):
            print('\nSTOP! Kraft_Filter is not null.\nDo this program manually.')
        SP.df_store = SP.df_store.drop(['Kraft_Filter'], axis=1)
    except:
        None

    # CLEAN 2: In the case where size A = 'Yes' and size B = 'YES'
    SP.df_item['Size'] = SP.df_item['Size'].str.upper()
    SP.df_store['BA_Filter'] = SP.df_store['BA_Filter'].str.upper()

    # CLEAN 3: Remove stores that are not part of the program
    SP.df_store = SP.df_store[SP.df_store['BA_Filter'] != 'NO']
    SP.df_store = SP.df_store.dropna()

    return SP

def preprocess(SP):
    # PREPROCESS 1: Calculate secondary capacity from capacity, casepack then drop columns
    SP.df_item['Sec_Cap'] = SP.df_item.apply(standard_sec_logic, axis=1)
    SP.df_item = SP.df_item.drop(['Capacity','CasePack'], axis=1)

    return SP

def produce_file_1(SP):
    # Inner join the stores and items on division and size where half is first/second/both and secondary capacity > 0.
    df_trans_in = pd.merge(SP.df_item.loc[(SP.df_item['Sec_Cap']>0) & (SP.df_item['Half']==1)], SP.df_store, how='inner', left_on=['Division','Size'], right_on=['LEGACY DIVISION','BA_Filter'])
    
    # drop columns not needed in transmission file
    df_trans_in = df_trans_in.drop(['Half','Size','Division','LEGACY DIVISION','BA_Filter'], axis=1)

    # add/rename columns needed in transmission file
    df_trans_in = df_trans_in.rename(columns={'STORE': 'STR_NUM', 'ItemNum': 'REF_NUM', 'Sec_Cap':'STR_MAINT_NUM'})
    df_trans_in['STR_ID'] = df_trans_in['STR_NUM']
    df_trans_in['STR_MAINT_TRANS_CD'] = 'IPSC'
    df_trans_in['STR_MAINT_CD'] = 'N'
    df_trans_in['STR_MAINT_CHAR_TXT'] = '(null)'
    df_trans_in['USER_ID'] = 'nuajd15'
    df_trans_in['Program_Name'] = SP.event_name

    # reorder columns
    df_trans_in = df_trans_in.reindex(columns = SP.trans_cols)

    # copy out file from in
    df_trans_out = df_trans_in.copy()
    df_trans_out['STR_MAINT_NUM'] = 0
            
    return [("IN", df_trans_in), ("OUT", df_trans_out)]
    
def produce_file_2(SP):
    SP.df_halves = load_half_data(SP.filename)

    # Inner join the stores and items on division and size where half is first/second/both and secondary capacity > 0.
    # first and both
    df_trans_first_in = pd.merge(SP.df_item.loc[(SP.df_item['Sec_Cap']>0) & ((SP.df_item['HalfWord']=='First')|(SP.df_item['HalfWord']=='Both'))], SP.df_store, how='inner', left_on=['Division','Size'], right_on=['Div','BA_Filter'])
    # first
    df_trans_first_out = pd.merge(SP.df_item.loc[(SP.df_item['Sec_Cap']>0) & (SP.df_item['HalfWord']=='First')], SP.df_store, how='inner', left_on=['Division','Size'], right_on=['Div','BA_Filter'])
    # second
    df_trans_second_in = pd.merge(SP.df_item.loc[(SP.df_item['Sec_Cap']>0) & (SP.df_item['HalfWord']=='Second')], SP.df_store, how='inner', left_on=['Division','Size'], right_on=['Div','BA_Filter'])
    # second and both
    df_trans_second_out = pd.merge(SP.df_item.loc[(SP.df_item['Sec_Cap']>0) & ((SP.df_item['HalfWord']=='Second')|(SP.df_item['HalfWord']=='Both'))], SP.df_store, how='inner', left_on=['Division','Size'], right_on=['Div','BA_Filter'])

    trans_list = [df_trans_first_in, df_trans_first_out, df_trans_second_in, df_trans_second_out]

    # columns that each df should have
    for i in trans_list:
        i.drop(['Half','Size','Division','Div','BA_Filter','HalfWord','Location','Quantity','Description'], axis=1, inplace=True)
        i.rename(columns={'StoreNum': 'STR_NUM', 'ItemNum': 'REF_NUM', 'Sec_Cap':'STR_MAINT_NUM'}, inplace=True)
        i.insert(0,'STR_MAINT_TRANS_ID','', allow_duplicates=True)
        i.insert(0,'STR_MAINT_CD','N', allow_duplicates=True)
        i.insert(0,'STR_MAINT_CHAR_TXT','(null)', allow_duplicates=True)
        i.insert(0,'PRCS_CD','T', allow_duplicates=True)
        i.insert(0,'USER_ID','nuajd15', allow_duplicates=True)
        i.insert(0,'CRT_TS','(null)', allow_duplicates=True)
        i.insert(0,'STR_ID', i['STR_NUM'], allow_duplicates=True)
        i.insert(0,'STR_MAINT_TRANS_CD','IPSC', allow_duplicates=True)
        i.insert(0,'Program_Name',SP.event_name, allow_duplicates=True)

    df_trans_first_in = df_trans_first_in.reindex(columns = SP.trans_cols)                   # what did this do?
    df_trans_first_out = df_trans_first_out.reindex(columns = SP.trans_cols)
    df_trans_second_in = df_trans_second_in.reindex(columns = SP.trans_cols)
    df_trans_second_out = df_trans_second_out.reindex(columns = SP.trans_cols)
                
    return [("FIRST IN", df_trans_first_in), ("FIRST OUT", df_trans_first_out), ("SECOND IN", df_trans_second_in), ("SECOND OUT", df_trans_second_out)]

def produce_file(SP):
    if SP.prog_size == 1:
        return produce_file_1(SP)
    elif SP.prog_size == 2:
        return produce_file_2(SP)
    else:
        print("Program size not 1 or 2")
        return None

def clean_file_name(str):
    # remove character windows can't use in a file name
    bad_chars = '<>:"/\\|?*'
    safe_name = ''.join(c for c in str if c not in bad_chars)
    return safe_name

def get_company_safe_folder():
    # Python says: "Can I write here?"
    # Windows: "...Hmm... fine, go ahead, it's just a small file."
    # Python is now inside the security circle for that folder and session.
    # So later, when Python says: "Now can I write a huge Excel file using a 3rd-party library?"
    # Windows says: "Okay. You're still the same process I let write a minute ago — carry on."
    
    # Try Downloads first
    try:
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        os.makedirs(downloads, exist_ok=True)
        test_file = os.path.join(downloads, "test_write.tmp")
        with open(test_file, 'w') as f:
            f.write("ok")
        os.remove(test_file)
        return downloads
    except Exception as e:
        print(f"Unable to make trusted folder: {e}")
        return None

def save_files(SP):
    # pick save folder
    # save_folder = get_company_safe_folder()
    # if save_folder == '':
    #     return None

    save_folder = r"Y:\QCP\Dept\Supply Chain\CAO\Bargain Aisle-Secondary Caps\2025\In and Out Files"

    # save an excel file for each transaction file
    for label_df_tup in SP.trans_files:
        label = label_df_tup[0]
        df = label_df_tup[1]
        try:
            file_name = clean_file_name(SP.event_name)
            SP.file_path = rf"{save_folder}\{file_name} {label}.xlsx"

            with pd.ExcelWriter(SP.file_path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False)

            print(print(f"File saved for {SP.event_name}: {SP.file_path}"))

        except Exception as e:
            print(f"Save Failed for {SP.event_name}: {e}")

## ---- MAIN ---- ##

SecProg = load_program()                            # select file, create class & dataframes from tables
if SecProg is not None:
    SecProg = clean(SecProg)                        # clean and apply secondary logic
    SecProg = preprocess(SecProg)                   # (*TO DO* Add sleeve code logic) Merge tables 
    SecProg.trans_files = produce_file(SecProg)     # create dataframes for CCAO upload files (in/out, one-part/two-halves)
    save_files(SecProg)                             # save as an excel file (***NOT WORKING***)

print("*" * 120)

# All data per program is saves to the SecProg class
# The transaction files are saved to SecProgClass.trans_files as a tuple: ("IN/OUT",)
