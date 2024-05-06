import pandas as pd
import time

path = r"E:\MIS\March\MAR24.csv"

df = pd.read_csv(path,usecols=['Date', 'Product_Description', 'Quantity', 'Unit', 'Indian_Exporter_Name', 'Foreign_Importer_Name', 'FOREIGN_COUNTRY', 'Indian_Port', 'Item_No'])

#DELETE DUPLICATE NO NEDD MARCH DATA NEXT TIME REQUIRED
df = df.drop_duplicates(subset=['Date', 'Product_Description', 'Quantity', 'Unit', 'Indian_Exporter_Name', 'Foreign_Importer_Name', 'Indian_Port', 'Item_No'])

# Drop rows containing "FROZEN" in the Product_Description column in place
df.drop(df[df["Product_Description"].str.contains(r'(?i)\bFROZEN\b', regex=True)].index, inplace=True)

# Add 'Shipment' column based on 'Item_No'
df['Shipment'] = df['Item_No'].apply(lambda x: 1 if x == 1 else None)

def filter(df, column_name, name):
    filtered_df = df[df[column_name].str.contains(fr'(?i)\b{name}\b', regex=True)]
    return filtered_df


print(df)