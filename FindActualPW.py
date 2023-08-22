import pandas as pd
import pyodbc

def find_actual_pw():
    # Load data from Excel into DataFrames
    needfind_df = pd.read_excel('file.xlsx', sheet_name='needfind')
    protocol_df = pd.read_excel('file.xlsx', sheet_name='protocol')
    fact_df = pd.read_excel('file.xlsx', sheet_name='fact')
    
    # Initialize BuyerCode column header in 'fact' DataFrame
    fact_df['BuyerCode'] = ""
    
    # Join 'fact' and 'needfind' DataFrames
    merged_df = pd.merge(needfind_df, fact_df, how='left', left_on=['Product', 'DistributorCode'], right_on=['Product', 'SellerCode'])
    
    # Initialize SQL connection
    conn_str = (
        r'DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};'
        r'DBQ=file.xlsx;'
    )
    cnxn = pyodbc.connect(conn_str)
    cursor = cnxn.cursor()
    
    # Section for s1
    sqlStr2 = "SELECT td.Product,td.FactoryCode,td.PWCode,ft.BuyerCode " \
              "FROM [s1$]td LEFT JOIN [fact$]ft ON td.PWCode =ft.SellerCode AND td.Product=ft.Product " \
              "GROUP BY td.Product,td.FactoryCode,td.PWCode,ft.BuyerCode"
    rst2 = cursor.execute(sqlStr2)
    s1_df = pd.DataFrame(rst2.fetchall(), columns=["Product", "FactoryCode", "PWCode", "SWCode"])
    
    # Perform matching and populating in a loop for s1 DataFrame
    
    # Section for s2
    sqlStr3 = "SELECT td.Product,td.FactoryCode,td.PWCode,td.SWCode,ft.BuyerCode " \
              "FROM [s2$]td LEFT JOIN [fact$]ft ON ft.SellerCode=td.SWCode AND td.Product=ft.Product " \
              "GROUP BY td.Product,td.FactoryCode,td.PWCode,td.SWCode,ft.BuyerCode"
    rst3 = cursor.execute(sqlStr3)
    s2_df = pd.DataFrame(rst3.fetchall(), columns=["Product", "FactoryCode", "PWCode", "SWCode", "TWCode"])
    
    # Perform matching and populating in a loop for s2 DataFrame
    
    # Section for s3
    sqlStr4 = "SELECT td.Product,td.FactoryCode,td.PWCode,td.SWCode,td.TWCode,ft.BuyerCode " \
              "FROM [s3$]td LEFT JOIN [fact$]ft ON ft.SellerCode=td.TWCode AND td.Product=ft.Product " \
              "GROUP BY td.Product,td.FactoryCode,td.PWCode,td.SWCode,td.TWCode,ft.BuyerCode"
    rst4 = cursor.execute(sqlStr4)
    s3_df = pd.DataFrame(rst4.fetchall(), columns=["Product", "FactoryCode", "PWCode", "SWCode", "TWCode", "QWCode"])
    
    # Perform matching and populating in a loop for s3 DataFrame

    
    # Save the resulting DataFrame to 'resultforfact' sheet in a new Excel file
    result_file = 'result.xlsx'
    merged_df.to_excel(result_file, sheet_name='resultforfact', index=False)

# Call the function to execute the code
find_actual_pw()
