import pysftp
import os
import paramiko
import pyodbc
import pandas as pd
import datetime
import sys

kiteworks_source_dir = 'ChangeHealth Data (SMS Response Data)'
local_destination_dir = 'aftercare_text_responses_dir'
hostname, port = 'secure.nychhc.org', 22
sftp_uid = ''
sftp_pwd = ''

proxy = paramiko.proxy.ProxyCommand('/usr/bin/nc --proxy healthproxy.health.dohmh.nycnet:8080 %s %d' % (hostname, port))
t = paramiko.Transport(sock=proxy)
t.connect(username=sftp_uid, password=sftp_pwd)
sftp = paramiko.SFTPClient.from_transport(t) # back to pysftp wrapper

print('Connected\n')
sftp.chdir(kiteworks_source_dir)

kw_fns = set()

for f in sftp.listdir():
    str_fn = f[0:13]
    if str_fn.lower() == 'aftercare sms':
        kw_fns.add(f)
    
print('Kiteworks filenames:', kw_fns, '\n')

# *Remember to authenticate with kinit: (1) Open Terminal window: JupyterHub home page > New > Other: Terminal; (2) kinit
driver= '{ODBC Driver 13 for SQL Server}'
server = 'SQLITICS1.health.dohmh.nycnet'
database = 'ICS_COVID19_TRACE_DATASHARE'
db_connect = pyodbc.connect(driver=driver, server=server, database=database, trusted_connection='yes', autocommit=True)
query1 = '''SELECT Filename
            FROM dap.aftercare_sms_responses'''
df_fns = pd.read_sql_query(query1, db_connect)

tbl_fns = set(df_fns.iloc[:,0])
print('SSMS filenames:', tbl_fns, '\n')

new_fns = kw_fns.difference(tbl_fns)
print('New filenames:', new_fns, '\n')

if len(new_fns) == 0:
    print('No new files.')
    sys.exit()

if not os.path.exists(local_destination_dir):
    os.mkdir(local_destination_dir)
    
database = 'ICS_COVID19_TRACE'
db_connect = pyodbc.connect(driver=driver, server=server, database=database, trusted_connection='yes')

list_f_dicts = []
for fn in new_fns:
    fn_path = os.path.join(local_destination_dir, fn)
    i_file_ext = fn.find('.')
    str_fn_date = fn[14:i_file_ext].strip()
    list_f_dicts.append({'filename': fn, 
                          'fn_path': fn_path, 
                          'fn_date': datetime.datetime.strptime(str_fn_date, '%m-%d-%Y').date()})
def get_fn_date(f_dict):
    return f_dict['fn_date']


list_f_dicts.sort(key=get_fn_date)

df_all = pd.DataFrame()
f_counter = 1
for f_dict in list_f_dicts:
    print(f'Starting iteration #{f_counter}.\n')
      
    sftp.get(f_dict['filename'], f_dict['fn_path'])
    print(f'Saved file #{f_counter}: "{f_dict["filename"]}" in folder "{local_destination_dir}".')
    
    #account for FILE FORMAT inconsistency across raw 'aftercare sms' files
    try:
        csv_data = pd.read_csv(f_dict['fn_path'], dtype = 'string')
    except:
        csv_data = pd.read_excel(f_dict['fn_path'])
        
    csv_data.columns = csv_data.columns.str.lower() #account for CASE inconsistency across raw 'aftercare sms' files
    
    #account for COLUMN HEADER inconsistency across raw 'aftercare sms' files
    try:
        df = csv_data[csv_data['contact result'] != 'Duplicate']
    except:
        df = csv_data[csv_data['call result'] != 'Duplicate']
       
    #account for COLUMN HEADER inconsistency across raw 'aftercare sms' files
    try:
        df['contact_date'] = csv_data['contact date']
    except:
        df['contact_date'] = csv_data['called time']
                
    
    #address COLUMN HEADER inconsistency across raw 'aftercare sms' files
    try:
        df['raw_id'] = csv_data['sales force id']
    except:
        df['raw_id'] = csv_data['patient id/appt. id']
    
    print(df.head())
    df = df[df['contact_date'] != '']
    
    max_nrows = 20000
    
    if len(df) <= max_nrows:
        list_dfs=[df]
        
    else:
        list_dfs = []
        for i in range(0, len(df)-max_nrows+1, max_nrows):
            list_dfs.append(df.iloc[i:i+max_nrows])
        
        remainder = len(df) % max_nrows
        if remainder > 0:
            list_dfs.append(df.iloc[len(df)-remainder:len(df)-1])
    
    print(f'Number of dfs in list_dfs: {len(list_dfs)}')
    
    df2 = pd.DataFrame(columns = ['Filename', 'Filename_Date', 'Contact_Date', 'Person_ID', 'Q_1A', 'Invalid_Response'])
    for df in list_dfs:
        print(f'Number of rows in this df: {len(df)}')
        id_strings = df['raw_id'].apply(lambda x: "'" + x + "'")
        id_strings.dropna(inplace=True)
        id_strings = list(set(id_strings))
        id_list = ', '.join(id_strings)
        
        query = '''SELECT SALESFORCE_ID, PARTY_EXTERNAL_ID
           FROM mv.COVID_MAVEN_TRACE_MATCH
           WHERE SALESFORCE_ID in (''' + id_list + ''')'''
        
        person_id_df = pd.read_sql_query(query, db_connect)
        person_id_dict = person_id_df.set_index('SALESFORCE_ID').T.to_dict('records')[0]
        
        temp_df = pd.DataFrame(columns = ['Filename', 'Filename_Date'])
        temp_df['Contact_Date'] = df['contact_date']
        temp_df['Person_ID'] = df['raw_id'].apply(lambda x: person_id_dict.get(x))          
        temp_df['Person_ID'].fillna('NA', inplace=True)
        temp_df['Filename'] = f_dict['filename']
        temp_df['Filename_Date'] = f_dict['fn_date']
        
        temp_df.insert(len(temp_df.columns), 'Q_1A', None)
        i=0
        for val in df['response']:
            try:
                temp_df.at[i, 'Q_1A'] = val.replace(' ', '').lower()
            except:
                pass
            i+=1
            
        temp_df['Q_1A'].replace(['yes', 'y'], int(1), inplace=True)
        temp_df['Q_1A'].replace(['no', 'n'], int(2), inplace=True)
        temp_df['Invalid_Response'] = df['invalid response']
        
        temp_df = temp_df[temp_df['Person_ID'] != 'NA']
        temp_df.fillna('NA', inplace=True)
        temp_df = temp_df[temp_df['Filename']!='NA']
        
        print('temp_df:\n',temp_df.head())
        print('temp_df info:\n', temp_df.info())
        
        df2 = df2.append(temp_df)
        df2.to_csv('df2_new.csv', index=False)
        
    df_all = df_all.append(df2)
    
    database = 'ICS_COVID19_TRACE_DATASHARE'
    db_connect = pyodbc.connect(driver=driver, server=server, database=database, trusted_connection='yes', autocommit=True)

    cursor = db_connect.cursor()
    
    for i in range(0, len(df2)):
        str_invalid = df2.iat[i,5].replace("'", "''")
        values = f"'{df2.iat[i,0]}', '{df2.iat[i,1]}', '{df2.iat[i,2]}', '{df2.iat[i,3]}', '{df2.iat[i,4]}', '{str_invalid}'"
        query3 = '''INSERT INTO dap.aftercare_sms_responses (Filename, Filename_Date, Contact_Date, Person_ID, Q_1A, Invalid_Response)
                VALUES (''' + values + ''')'''
        cursor.execute(query3)
        db_connect.commit()
    
    print(f"Finished inserting rows from file {f_dict['filename']}\n")
    
    f_counter+=1

df_all.to_csv('df_save_new_2.csv', index=False)

sftp.close()
