import pandas as pd
import re
import openpyxl
import xlwings
import tkinter as tk
from tkinter import filedialog
import os
import xlsxwriter
from xlsxwriter import styles


# Read CSV file
df = pd.read_csv(r"E:\MIS\March\MAR24.CSV", usecols=['Date', 'Product_Description', 'Quantity', 'Unit', 'Indian_Exporter_Name', 'Foreign_Importer_Name', 'FOREIGN_COUNTRY', 'Indian_Port', 'Item_No'])

print(type(df['Quantity']))

#DELETE DUPLICATE NO NEDD MARCH DATA NEXT TIME REQUIRED
df = df.drop_duplicates(subset=['Date', 'Product_Description', 'Quantity', 'Unit', 'Indian_Exporter_Name', 'Foreign_Importer_Name', 'Indian_Port', 'Item_No'])

# Drop rows containing "FROZEN" in the Product_Description column in place
df.drop(df[df["Product_Description"].str.contains(r'(?i)\bFROZEN\b', regex=True)].index, inplace=True)

# Add 'Shipment' column based on 'Item_No'
df['Shipment'] = df['Item_No'].apply(lambda x: 1 if x == 1 else None)

def process_data_importer(df):
    # Group by Indian_Exporter_Name and aggregate Quantity and Shipment
    grpup = df.groupby(by=["Indian_Exporter_Name",'Unit','FOREIGN_COUNTRY']).aggregate({"Quantity": "sum", "Shipment": "sum"}).reset_index()

    # Pivot the data
    pivot = df.pivot_table(index="Indian_Exporter_Name", columns="Product_Description", values='Quantity', aggfunc="sum", fill_value=None).reset_index()
    pivot = pivot.rename(columns={'': "Indian_Exporter_Name"})

    # Merge the grouped data and pivot table
    merged_df = grpup.merge(pivot, how='left', on="Indian_Exporter_Name")
    merged_df['Sr No'] = merged_df.index + 1
    merged_df.set_index('Sr No', inplace=True)

    return merged_df

def process_data_exporter(df):
    # Group by Indian_Exporter_Name and aggregate Quantity and Shipment
    # grpup = df.groupby(by=['Foreign_Importer_Name','Unit','FOREIGN_COUNTRY']).aggregate({"Quantity": "sum", "Shipment": "sum"}).reset_index()

    # Pivot the data
    pivot = df.pivot_table(index=['Foreign_Importer_Name','Unit','FOREIGN_COUNTRY'], columns="Product_Description", values=['Quantity','Shipment'], aggfunc="sum", fill_value=None,margins=True,margins_name='Total')
    # pivot = pivot.rename(columns={'': 'Foreign_Importer_Name'})

    # # Merge the grouped data and pivot table
    # merged_df = grpup.merge(pivot, how='left', on='Foreign_Importer_Name')
    # merged_df['Sr No'] = merged_df.index + 1
    # merged_df.set_index('Sr No', inplace=True)

    return pivot

def filter(df, column_name, name):
    filtered_df = df[df[column_name].str.contains(fr'(?i)\b{name}\b', regex=True)]
    return filtered_df
  
# Standardize 'Product_Description'
#MIX PRODUCT
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bDRY POTATO\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bROASTED SALTED\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bCARGO VEGETABLES\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bFRESH VEGETABLES\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bindian mixed\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\brosemary\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bTHYME\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bTurmeric\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bmelon|MELONS\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bWATERMELON|WATERMELONS\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bWATERMIXED\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bAMLA\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bAPPLE|APPLE BER|ber\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bARVI\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bBABY POTATO\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bBANANA\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bcurry|curry leave|curry leaves\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bDRUM|DRUSMSTIC|DRUMSTICKS|DRUM STICK\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bbCARROT|PUMPKIN|CARROTS|BIITER|MIXED|MIXEDS|LEAVES|PEA|GREEN PEA|GALKA|GAWAR|GREEN BEANS|PAPAYA|GREEN BITTER\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bSAPOTA|SINGHARA|SINGODA|SUGARCANE|SURAN|TINDA|TINDORI|TURIYA|VALLERI|VALORE|SUGAR|RED CARROT|RED PUMKIN\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bMIXED\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bMIXEDSTICK\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bMIXEDSTICKS\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bMIXEDSTICKS,\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'CHICKOO.*|FRESH RED BORE.*|FRESH TONDORI.*|FRESH VAL PAPADI.*|GUAVA.*', 'MIXED', regex=True)

df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bPARWAL\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bPAPDI\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bVAL PAPDI,\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bBORE,\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bred BORE,\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bTANDORI,\b.*', 'MIXED', regex=True)

df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bLEMON\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bKAND\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bKOLA\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bLOTUS\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bPARWAR\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bMIXEDS FRESH\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bLONG PADWAL\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bPINEAPPLE\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bPOTATOS\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bPOTATO\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bKIMAYE MIX CUP\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bTHOMPSON\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bMIX VEGETABLES\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bBAEL\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bMIX VEGETABLES\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bMIX FRUITS\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bCOCOBULK,2KG-EXP\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bMIX VEGITABLE\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bMIX VEGITABLE\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bRAISIN\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bBOR\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bCARROT\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bBEANS\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bGARLIC\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bCHIKOO\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bMIX LENTILS\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bAWLA\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bGRAGON\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bRAMFAL\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bSTAR FRUIT\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bASSORTED LENTILS\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bFRESH MIXED\b.*', 'MIXED', regex=True)

df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bNotify1: SUCCESS TYCOON LTD. Adress: UNIT 1 ,16TH FLOOR , YUEN LONG TRADING CENTRE,33 WANG YIP STREET WEST ,YUEN LONG .N\b.*', 'MIXED', regex=True)

#GRAPES ALSO IN MIX PRODUCT
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bGRAPES\b.*', 'MIXED', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bGRAPE\b.*', 'MIXED', regex=True)

#POMOGRANATE PRODUCT
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bFRESH POMEGRANATE GROSS.\b.*', 'POMEGRANATE', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bPOMEGRANATEWT:3.50/PAC(PACKED IN CORRUGATED BOXES NET WT. 3.0)\b.*', 'POMEGRANATE', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bPOMEGRANATE GROSS.WT:4.0PAC (PACKED IN CORRUGATED BOXES NET WT 3.5KGS).\b.*', 'POMEGRANATE', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bPOMEGRANATES GROSS WT:4.50/PAC(PACKED INCORRUGATED BOXES NET WT. 4.0)\b.*', 'POMEGRANATE', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bPOMEGRANATES\b.*', 'POMEGRANATE', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bPOMEGRANATEWT:3.50/PAC(PACKED IN CORRUGATED BOXES NET WT. 3.0)\b.*', 'POMEGRANATE', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bPOMEGRANATE\b.*', 'POMEGRANATE', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bANAR\b.*', 'POMEGRANATE', regex=True)

#POM ARILS
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bPOMEGRANATE ARILS\b.*', 'POMEGRANATE ARILS', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bFRESH POMEGRANATE  ARILS\b.*', 'POMEGRANATE ARILS', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bARILS cup\b.*', 'POMEGRANATE ARILS', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bARILSpunnet\b.*', 'POMEGRANATE ARILS', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bARILS\b.*', 'POMEGRANATE ARILS', regex=True)

#COCONUT PRODUCT
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bCOCONUT\b.*', 'COCONUTS', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bCOCONUTS\b.*', 'COCONUTS', regex=True)

#BABY CORN PRODUCT
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bBABY CORN\b.*', 'Baby Corns', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bBABY CORNS\b.*', 'Baby Corns', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bBABYCORNS\b.*', 'Baby Corns', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bBABYCORN\b.*', 'Baby Corns', regex=True)

#DUDHI OR OKRA
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bDUDHI\b.*', 'DUDHI', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bBOTTLE GAURD\b.*', 'DUDHI', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bBOTTLE GOURD\b.*', 'DUDHI', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bBOTTLE GUARD GROSS\b.*', 'DUDHI', regex=True)

#OKRA
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bOKRA\b.*', 'OKRA', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bOKRAS\b.*', 'OKRA', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bFRESH OKRA\b.*', 'OKRA', regex=True)

#CHILLI PRODUCT
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bCHILLI\b.*', 'CHILLI', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bCHILLY\b.*', 'CHILLI', regex=True)

#MANGO PRODUCT
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bMANGO\b.*', 'MANGO', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bMANGOES\b.*', 'MANGO', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bALPHONSO,|ALPHONSO|BANGANAPALLI\b.*', 'MANGO', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bkesar,|KESAR, 12\b.*', 'MANGO', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bMANGO 12 PC,\b.*', 'MANGO', regex=True)

#GUVAVA PRODUCT
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bWHITE GUAVA\b.*', 'GUAVA', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bGUAVA 3.5 KG\b.*', 'GUAVA', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bGUAVA\b.*', 'GUAVA', regex=True)

#DRAGON FRUIT PRODUCT
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bDRAGON FRUIT\b.*', 'DRAGON FRUIT', regex=True)
df["Product_Description"]=df["Product_Description"].str.replace(r'(?i).*\bDRAGONFRUIT\b.*', 'DRAGON FRUIT', regex=True)


#Standardize 'Indian_Exporter_Name'
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(KAY BEE|Kay Bee)\b.*', 'Kay Bee Exports', regex=True)
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(NAMDHARI| NAMDHARI)\b.*', 'NAMDHARI SEEDS PVT LTD', regex=True)
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(NEXTON|NEXTON FOODS)\b.*', 'NEXTON FOODS', regex=True)
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(GREEN AGREVOLUTION|GREEN AGREVOLUTION PRIVATE LIMITED)\b.*', 'GREEN AGREVOLUTION', regex=True)
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(BARAMATI|BARAMATI AGRO)\b.*', 'BARAMATI AGRO LIMITED', regex=True)
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(ULINK|ULINK AGRITECH)\b.*', 'ULINK AGRITECH', regex=True)
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(VASHINI EXPORTS|VASHINI)\b.*', 'VASHINI EXPORTS', regex=True)
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(KHUSHI INTERNATIONAL)\b.*', 'KHUSHI INTERNATIONAL', regex=True)
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(SIA IMPEX|SIA IM)\b.*', 'SIA IMPEX', regex=True)
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(IFC|IFC OVERSEAS )\b.*', 'IFC OVERSEAS', regex=True)
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(MANTRA INTERNATIONAL|MANTRA)\b.*', 'MANTRA INTERNATIONAL', regex=True)
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(SCION|SCION AGRICOS)\b.*', 'SCION INTERNATIONAL', regex=True)
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(SUPER FRESH|SUPER F)\b.*', 'SUPER FRESH', regex=True)
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(M. K.|M. K. EXPORTS)\b.*', 'M.K.EXPORTS', regex=True)
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(ALL SEASON|ALL SEASON EXPORTS)\b.*', 'ALL SEASON EXPORTS', regex=True)
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(THREE CIRCLES|THREE CIRCLES AGRO)\b.*', 'THREE CIRCLES AGRO', regex=True)
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(GO GREEN|GO GREEN EXPORTS)\b.*', 'GO GREEN EXPORTS', regex=True)
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(KASHI EXPORTS)\b.*', 'KASHI EXPORTS', regex=True)
df['Indian_Exporter_Name'] = df['Indian_Exporter_Name'].str.replace(r'(?i).*\b(MAGNUS|MAGNUS FARM)\b.*', 'MAGNUS FARM', regex=True)


#Standardize 'Foreign_Importer_Name'
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(TO ORDER|TO THE|TO THE ORDER)\b.*', 'TO ORDER', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace('-', 'TO ORDER')
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace('TO', 'TO ORDER')
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(',', 'TO ORDER')
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'\.{4,}|\.', 'TO ORDER', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'\_{4,}', 'TO ORDER', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(Flamingo|FLAMINGO)\b.*', 'FLAMINGO PRODUCE', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(WEALMOOR|WEAL MOOR)\b.*', 'WEALMOOR LTD', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(MINOR|MINOR,)\b.*', 'MINOR WEIR & WILLIS LTD', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(Yukon|YUKON INTERNATION,)\b.*', 'Yukon', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(RAJA FOODS L|RAJA FOOD L,)\b.*', 'RAJA FOODS LLC', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(BARAKAT|BARAKAT VEGETABLES,)\b.*', 'BARAKAT VEGETABLES', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(BARFOOTS|BARFOOTS OF|BARFOOT,)\b.*', 'Barfoots Of Botley Ltd', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(COREFRESH|COREFRESH LTD|COREFRESH LIMITED,)\b.*', 'COREFRESH LTD', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(KAY BEE|Kay Bee Veg)\b.*', 'Kay Bee Exports', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(NRTC DUBAI|NRTC DUBAI INTERNATIONAL)\b.*', 'NRTC DUBAI INTERNATIONAL', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(S & F GLOBAL|S&F GLOBAL)\b.*', 'S&F GLOBAL FRESH EXOTICS LTD', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(ali gholami|ALI GHOLAMI VEGETABLES| ALI GHOLAMI)\b.*', 'ALI GHOLAMI VEGETABLES', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(VIVA ENTERPRISE|viva enter)\b.*', 'VIVA ENTERPRISE', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(CG FOODS|C.G. FOODS|C.G.FOODS)\b.*', 'C.G FOODS', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(KOLLA OVERSEAS|KOLLA)\b.*', 'KOLLA OVERSEAS', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(PAPLU|PAPLU FRESH VEG)\b.*', 'KOLLA OVERSEAS', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(Limited Liability|Limited Liability Company)\b.*', 'Limited Liability Company', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(PRIDE J.V|Nature Pride)\b.*', 'NATURE PRIDE', regex=True)
df['Foreign_Importer_Name'] = df['Foreign_Importer_Name'].str.replace(r'(?i).*\b(CTO ORDERGTO)\b.*', 'CTO ORDERGTO', regex=True)


# Convert all text columns to uppercase
df = df.applymap(lambda x: x.upper() if isinstance(x, str) else x)




#PRODUCT REPORT
# baby_report = df[df['Product_Description'].str.contains(r'(?i)\b(BABY CORNS)\b')]

# Filter rows where 'Foreign_Importer_Name'
flamingo = filter(df,'Foreign_Importer_Name',"flamingo")
wealmoor = filter(df,'Foreign_Importer_Name',"wealmoor")
mww = filter(df,'Foreign_Importer_Name',"minor")
nature = filter(df,'Foreign_Importer_Name',"NATURE PRIDE")
rajafood = filter(df,'Foreign_Importer_Name',"RAJA FOODS")
yukon = filter(df,'Foreign_Importer_Name',"yukon")


# Filter rows where 'Indian_Exporter_Name' is '
magnus = filter(df,'Indian_Exporter_Name','MAGNUS FARM')
namdhari = filter(df,'Indian_Exporter_Name','NAMDHARI')
kb = filter(df,'Indian_Exporter_Name','KAY BEE')
dehat = filter(df,'Indian_Exporter_Name','GREEN AGREVOLUTION')
vashini = filter(df,'Indian_Exporter_Name','VASHINI EXPORTS')
baramati = filter(df,'Indian_Exporter_Name','BARAMATI AGRO LIMITED')
ulink = filter(df,'Indian_Exporter_Name','ULINK AGRITECH')
kashi_export = filter(df,'Indian_Exporter_Name','KASHI EXPORTS')
khushi =filter(df,'Indian_Exporter_Name','KHUSHI INTERNATIONAL')
go_green = filter(df,'Indian_Exporter_Name','GO GREEN EXPORTS')
three_circle =filter(df,'Indian_Exporter_Name','THREE CIRCLES AGRO')
all_season = filter(df,'Indian_Exporter_Name','ALL SEASON EXPORTS')
mk = filter(df,'Indian_Exporter_Name','M.K.EXPORTS')




# Create pivot table for importer
flamingo_pivot = process_data_importer(flamingo)
wealmoor_pivot = process_data_importer(wealmoor)
mww_pivot = process_data_importer(mww)
nature_pivot = process_data_importer(nature)
yukon_pivot = process_data_importer(yukon)
rajafood_pivot = process_data_importer(rajafood)


# #create pivot for indian exporter
magnus_pivot = process_data_exporter(magnus)
namdhari_pivot = process_data_exporter(namdhari)
kb_pivot = process_data_exporter(kb)
dehat_pivot = process_data_exporter(dehat)
vashini_pivot = process_data_exporter(vashini)
baramati_pivot = process_data_exporter(baramati)
ulink_pivot = process_data_exporter(ulink)
kashi_pivot = process_data_exporter(kashi_export)
khushi_pivot = process_data_exporter(khushi)
go_green_pivot = process_data_exporter(go_green)
three_circle_pivot = process_data_exporter(three_circle)
all_season_pivot = process_data_exporter(all_season)
mk_pivot = process_data_exporter(mk)


# Save pivot table to Excel Importer
with pd.ExcelWriter(r'D:\UserProfile\Desktop\MIS\March\FOREIGN_IMPORTER_REPORT.xlsx',engine="xlsxwriter") as writer:
    flamingo_pivot.to_excel(writer,sheet_name='Flamingo')
    wealmoor_pivot.to_excel(writer,sheet_name='Wealmoor')
    mww_pivot.to_excel(writer,sheet_name='MWW')
    nature_pivot.to_excel(writer,sheet_name='Nature Pride')
    yukon_pivot.to_excel(writer,sheet_name='Yukoon')
    rajafood_pivot.to_excel(writer,sheet_name='Rajafood')
    
# #indian_exporter_report_export_in_excel
with pd.ExcelWriter(r'D:\UserProfile\Desktop\MIS\March\INDIAN_EXPORTER_REPORT.xlsx',engine="xlsxwriter") as writer:
    namdhari_pivot.to_excel(writer,sheet_name='NAMDHARI')
    kb_pivot.to_excel(writer,sheet_name='KAY BEE')
    dehat_pivot.to_excel(writer,sheet_name="DEHAT")
    vashini_pivot.to_excel(writer,sheet_name="vashini")
    magnus_pivot.to_excel(writer,sheet_name="MAGNUS")
    baramati_pivot.to_excel(writer,sheet_name="BARAMATI")
    ulink_pivot.to_excel(writer,sheet_name="ULINK")
    kashi_pivot.to_excel(writer,sheet_name="KASHI EXPORTS")
    khushi_pivot.to_excel(writer,sheet_name="KHUSHI INTERNATIONAL")
    go_green_pivot.to_excel(writer,sheet_name="GO GREEN")
    three_circle_pivot.to_excel(writer,sheet_name="THREE CIRCLES")
    all_season_pivot.to_excel(writer,sheet_name="ALL SEASONS")
    mk_pivot.to_excel(writer,sheet_name="MK EXPORTS")
    

