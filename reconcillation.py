import pandas as pd
# Import Excel file
df_crf = pd.read_excel('C:/Users/ZYAH/OneDrive - Novo Nordisk/Azhar/Python/4909-DMMP-V.0.1.xlsx', sheet_name='CRF - Lab data')
subjid=df_crf['USUBJID'].str.slice(-6,None)
df_crf['subjid']=subjid.astype('object')
columns_to_extract = ['subjid','LBDTC']
df1=df_crf[columns_to_extract].copy()

df_ext=pd.read_excel('C:/Users/ZYAH/OneDrive - Novo Nordisk/Azhar/Python/4909-DMMP-V.0.1.xlsx', sheet_name='External lab')
df_ext.rename(columns={'SUBJID': 'subjid'}, inplace=True)
df2=df_ext[columns_to_extract].copy()
df2['subjid']=df2['subjid'].astype(object)
#print(df1.dtypes)
#print(df2.dtypes)
df3 = pd.merge(df2,df1, on=['subjid'], how='right', indicator=True)
print(df3.head(100))
path="C:/Users/ZYAH/OneDrive - Novo Nordisk/Azhar/Python/output.xlsx"
df3.to_excel(path,index=False)


#print(df_ext.columns)
# Merge the two DataFrames based on the common column
#merged_df = pd.merge(df_crf, df_ext, on='subjid', how='outer', suffixes=('_crf', '_ext'))

#print(merged_df.columns)
#for subjid_value in df_crf['subjid']:
    #print(subjid_value)
    #filter_1=df_crf['subjid']==df_ext['subjid']
    #filter_2=df_crf['LBDTC']==df_ext['LBDTC']
    

#print(filter_df.head())
# Export to Excel file
#df.to_excel('new_filename.xlsx', index=False