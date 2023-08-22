import pandas as pd
import pyodbc

def find_protocol_pw():
    # Load data from Excel into DataFrames
    needfind_df = pd.read_excel('file.xlsx', sheet_name='needfind')
    protocol_df = pd.read_excel('file.xlsx', sheet_name='protocol')
    resultforprotocol_df = pd.read_excel('file.xlsx', sheet_name='resultforprotocol')
    tempdata_df = pd.read_excel('file.xlsx', sheet_name='tempdata')
    
    # Calculate total rows in various sheets
    needfindTotalRow = needfind_df.shape[0]
    protocolTotalRow = protocol_df.shape[0]
    resultforprotocolTotalRow = resultforprotocol_df.shape[0]
    tempdataTotalRow = tempdata_df.shape[0]
    
    # Add "Order" column to "needfind" DataFrame
    needfind_df['Order'] = needfind_df.index + 2
    
    # Set up connection to Excel
    conn_str = (
        r'DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};'
        r'DBQ=file.xlsx;'
    )
    cnxn = pyodbc.connect(conn_str)
    cursor = cnxn.cursor()
    
    # SQL query and data retrieval for resultforprotocol sheet
    sqlStr = "SELECT nf.Product,nf.DistributorCode,nf.DistributorName,nf.DistributorLevel,p.UpperDistributorCode,p.DistributorLevel,null,null,nf.Order " \
             "FROM [needfind$]nf LEFT JOIN [protocol$]p ON nf.Product = p.Product AND nf.DistributorCode = p.DistributorCode ORDER BY nf.[order]"
    rst = cursor.execute(sqlStr).fetchall()
    resultforprotocol_df = pd.DataFrame(rst, columns=["Product", "DistributorCode", "DistributorName", "DistributorLevel", "UpperDistributorCode", "ProductProtocolDistributorLevel", "PrimaryDistributorCode", "PrimaryWholesalerName", "Order"])

    # Loop to update data in resultforprotocol DataFrame based on conditions
    for i in range(1, resultforprotocolTotalRow):
        if resultforprotocol_df.loc[i, 'UpperDistributorCode'] == "Roche" or resultforprotocol_df.loc[i, 'UpperDistributorCode'] == "D00010893":
            resultforprotocol_df.loc[i, 'PrimaryDistributorCode'] = resultforprotocol_df.loc[i, 'DistributorCode']
            resultforprotocol_df.loc[i, 'PrimaryWholesalerName'] = resultforprotocol_df.loc[i, 'DistributorName']
        elif pd.isna(resultforprotocol_df.loc[i, 'UpperDistributorCode']):
            resultforprotocol_df.loc[i, 'PrimaryDistributorCode'] = "Cannot Find"
    
    # SQL query and data retrieval for tempdata DataFrame
    sqlStr = "SELECT Product,DistributorCode,UpperDistributorCode,PrimaryDistributorCode,null,Order " \
             "FROM [resultforprotocol$] " \
             "WHERE PrimaryDistributorCode is null"
    rst = cursor.execute(sqlStr).fetchall()
    tempdata_df = pd.DataFrame(rst, columns=["Product", "DistributorCode", "UpperDistributorCode", "PrimaryDistributorCode", "Factory", "Order"])
    
    # Loop to update data in tempdata DataFrame
    for i in range(tempdataTotalRow):
        tempdata_df.loc[i, 'PrimaryDistributorCode'] = tempdata_df.loc[i, 'DistributorCode'] + tempdata_df.loc[i, 'UpperDistributorCode']
        tempdata_df.loc[i, 'Factory'] = protocol_df.loc[protocol_df['ProductDistributorCode'] == tempdata_df.loc[i, 'PrimaryDistributorCode'], 'PrimaryWholesalerName'].values[0]
    
    # Loop to update data in resultforprotocol DataFrame
    for i in range(1, resultforprotocolTotalRow):
        if pd.isna(resultforprotocol_df.loc[i, 'PrimaryWholesalerName']):
            resultforprotocol_df.loc[i, 'PrimaryWholesalerName'] = protocol_df.loc[protocol_df['ProductDistributorCode'] == resultforprotocol_df.loc[i, 'PrimaryDistributorCode'], 'PrimaryWholesalerName'].values[0]
    
    # Save the resulting DataFrames back to Excel file
    resultforprotocol_df.to_excel('file.xlsx', sheet_name='resultforprotocol', index=False)
    tempdata_df.to_excel('file.xlsx', sheet_name='tempdata', index=False)
    
    # Delete unnecessary columns
    del resultforprotocol_df['ProductProtocolDistributorLevel']
    del resultforprotocol_df['UpperDistributorCode']
    del resultforprotocol_df['DistributorName']
    del resultforprotocol_df['DistributorCode']
    del resultforprotocol_df['DistributorLevel']
    tempdata_df.drop(['DistributorCode', 'UpperDistributorCode'], axis=1, inplace=True)

# Call the function to execute the code
find_pw()
