#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import os
import glob
import datetime
import pyodbc

home = "/run/user/1701720350/gvfs/smb-share:server=gcsecshare100,share=share"


# In[2]:


# Create a dictionary mapping extract filename roots to SQL database tables
tbl_dict = {'Interaction': 'interactions', 
            'Result': 'results',
            'Contact': 'contacts'}


# In[3]:


def readFile(file, ID, input_cols, fields):
    csv_cols = pd.read_csv(file, dtype = 'string').columns
    missing_cols = list(set(input_cols).difference(set(csv_cols)))
    
    if fields != 'all':
        valid_cols = list(set(input_cols).intersection(set(csv_cols)))
    elif fields == 'all':
        valid_cols = csv_cols
    
    df_file = pd.read_csv(file, usecols = valid_cols, dtype = 'string')
    df_file.set_index(ID, drop = False, inplace = True)
    df_file = df_file
    
    # If the file does not contain a specificed column, fill with 'Missing value'
    for c in missing_cols:
        df_file = df_file.assign(**{c: "Missing value"})
        
    return df_file


# In[4]:


def compileExtracts(tbl, ID, file_start_dt, fields):
    
    tblRoot = tbl_dict[tbl]
    
    if datetime.datetime.strptime(file_start_dt, '%Y-%m-%d').date() < datetime.date(2021, 2, 17):
        file_start_dt = '2021-02-17'
       
    baseFile = home+'/Trace/Inbound/processed/'+tblRoot+'_'+file_start_dt+'.csv'
    print('Reading extract files...')
    
    if fields != 'all':
        baseCols = fields
        baseCols.insert(0, 'LastModifiedDate')
        baseCols.insert(0, ID)
    
    elif fields == 'all':
        baseCols = pd.read_csv(baseFile, dtype = 'string').columns
    
    df = readFile(baseFile, ID, baseCols, fields)
    print('Base file "' + os.path.basename(baseFile) + '" read.')
    
    files = []
    
    for f in glob.glob(home+'/Trace/Inbound/processed/'+tblRoot+'_2021*.csv'):
        date = datetime.datetime.strptime(os.path.basename(f)[-14:-4], '%Y-%m-%d').date()
        if date > datetime.datetime.strptime(file_start_dt, '%Y-%m-%d').date(): 
            files.append(f)
    
    for f in files:
        df_temp = readFile(f, ID, df.columns, fields)
        
        # Concatenate dataframes and drop first instance of any duplicates (i.e. add new records and update existing records)
        df_concat = pd.concat([df, df_temp])
        df = df_concat[~df_concat.index.duplicated(keep='last')]
        print('File "' + os.path.basename(f) + '" processed.')
        
    print('File processing completed.')
    
    df = df.copy()
    df['Date'] = df['LastModifiedDate'].apply(lambda x: datetime.datetime.strptime(x[0:10], '%Y-%m-%d').date())

    df.fillna('Missing value', inplace=True)
            
    getMonth = lambda x: x.strftime('%b')
    col = df.Date.apply(getMonth)
    df = df.assign(Month_last_modified = col.values)
           
    print('Head: ', df.head())
    print('Tail: ', df.tail())
    
    return df


# In[5]:


# Function to check a table's extracts for non-null/non-blank values by field
def checkExtracts(tbl, df_extracts, str_start_dt, df_summary):
    
    df = df_extracts
    
    fields = list(df.columns)
    fields.remove('Date')
    fields.remove('Month_last_modified')
    
    start_date = datetime.datetime.strptime(str_start_dt, '%Y-%m-%d').date()  
    df = df[df['Date'] >= start_date]
    print(df.info())
    
    headers = list(df_summary.columns)
    headers.remove('Table')
    headers.remove('Field Name')
    
    for h in headers: df_summary = df_summary.rename(columns = {h: h+' Extract'})
    
    # Fill the template
    row=0
    for field in fields:
        df_summary.at[row, 'Table'] = tbl
        df_summary.at[row, 'Field Name'] = field
               
        c = pd.crosstab(df[field], df.Month_last_modified, margins = True, margins_name = 'Overall')
 
        for header in headers: 
            try:
                total = c.loc['Overall', header]
                try:
                    df_summary.at[row, header + ' Extract'] = c.loc['Overall', header] - c.loc['Missing value', header]
                except:
                    df_summary.at[row, header + ' Extract'] = c.loc['Overall', header]
            except:      
                pass
            
        print('Extract population of field "' + field + '" completed.')

        row+=1
        
    return df_summary


# In[6]:


def checkSQL(tbl, start, df_summary, fields, ID, extracts_on):
    
    # *Remember to authenticate with kinit: (1) Open Terminal window: JupyterHub home page > New > Other: Terminal; (2) kinit
    driver= '{ODBC Driver 13 for SQL Server}'
    server = 'SQLITICS1.health.dohmh.nycnet'
    database = 'ICS_COVID19_TRACE_DATASHARE'
    
    db_connect = pyodbc.connect(driver=driver, server=server, database=database, trusted_connection='yes')

    headers = list(df_summary.columns)
    headers.remove('Table')
    headers.remove('Field Name')
    
    #FIELDS
    if not extracts_on: 
  
        if fields == 'all':
            query1 = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'SF2' AND TABLE_NAME = '"+ tbl + "'"
            column_names = pd.read_sql_query(query1, db_connect)
            fields = column_names['COLUMN_NAME']
            
                    
        elif fields != 'all':
            fields.insert(0, 'LastModifiedDate')
            fields.insert(0, ID)
        
        row=0
        for field in fields:
            df_summary.at[row, 'Table'] = tbl
            df_summary.at[row, 'Field Name'] = field
            row+=1
            
        for header in headers: df_summary = df_summary.rename(columns = {header: header+' SQL'})
    
    
    df_summary = df_summary.set_index('Field Name', drop = False)
    

    if extracts_on: 
        headers = list(map(lambda x: x[:-8], headers))
        fields = list(df_summary.index)
        
        i=3
        for header in headers:
            df_summary.insert(i, header+' SQL', [None]*len(df_summary.index))
            i+=2           
    
    startdt = '\'' + start + '\''     
    
    print('Populating field rows with SQL query results...')
    for field in fields:
        query = '''SELECT
                    COALESCE(FORMAT(CAST(LastModifiedDate as datetime), 'MMM'), 'Overall') AS Month_last_modified,
                    COUNT(''' + field + ''') AS Non_NULL_Count
                    FROM ICS_COVID19_TRACE_DATASHARE.SF2.''' + tbl + '''
                    WHERE (''' + field + ''' IS NOT NULL 
                        AND ''' + field + ''' <> ''
                        AND CAST(LastModifiedDate as datetime) >= ''' + startdt + ''')
                    GROUP BY FORMAT(CAST(LastModifiedDate as datetime), 'MMM')
                    WITH ROLLUP'''
               
        try: 
            d = pd.read_sql_query(query, db_connect).set_index('Month_last_modified')
        except:
            print('Failed to process field ' + field)

        
        for header in headers:
            try:
                 df_summary.at[field, header + ' SQL'] = d.loc[header, 'Non_NULL_Count']
            except:
                print('Failed to process field ' + field + ' for header ' + header)
        
        print('SQL population of field "' + field + '" completed.')

    print('SQL check completed.')
    
    return df_summary


# In[7]:


def checkVariance(df_summary):
  
    df_summary.fillna(0, inplace = True)   
    fields = df_summary.index
    
    headers = list(df_summary.columns)
    headers.remove('Table')
    headers.remove('Field Name')
    
    varheaders = []
    i=0
    while i < len(headers):
        varheaders.append(headers[i][:-8])
        i+=2
        
    i=0
    j=4
    while i < len(varheaders):
        df_summary.insert(j, varheaders[i]+' Variance', [None]*len(df_summary.index))
        i+=1
        j+=3
    
    for field in fields:
        for header in varheaders:
            df_summary.at[field, header + ' Variance'] = df_summary.loc[field, header + ' Extract'] - df_summary.loc[field, header + ' SQL']

    return df_summary


# In[8]:


def checkTable(table, ID, fields, extract_file_strt_dt, modified_strt_dt, months, extracts, sql, write_to_filename):
    
    columns = ['Table', 'Field Name', 'Overall']
    columns.extend(months)
    summary_template = pd.DataFrame(columns = columns)
    df_extracts = pd.DataFrame()
    
    if not extracts and not sql:
        print('Please set "extracts" and/or "sql" to true.')
        return None, None
    
    if extracts:
        df_extracts = compileExtracts(table, ID, extract_file_strt_dt, fields)
        extracts_summary = checkExtracts(table, df_extracts, modified_strt_dt, summary_template)
        print(extracts_summary.head())
        
        if not sql:
            tbl_summary = extracts_summary
            
        else:
            sql_summary = checkSQL(table, modified_strt_dt, extracts_summary, fields, ID, extracts)
            tbl_summary = checkVariance(sql_summary)
            
    if not extracts:
        tbl_summary = checkSQL(table, modified_strt_dt, summary_template, fields, ID, extracts)
    
    tbl_summary.to_csv(write_to_filename, index=False)
    
    print('Table check completed. Results saved to "' + write_to_filename + '".')
    
    return df_extracts, tbl_summary


# In[12]:

# Example run below:

parameters = {
    'table': 'Interaction',
    'ID': 'Name',
    'fields': 'all',
    'extract_file_strt_dt': '2021-08-24',
    'modified_strt_dt': '2021-08-24',
    'months': ['Aug', 'Sep'],
    'extracts': True,
    'sql': True,
    'write_to_filename': 'int_summary_2021-09-24.csv'
}

int_extracts, int_summary = checkTable(**parameters)
