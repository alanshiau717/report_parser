import pandas as pd
import os
import re
from env import PATH



def dir_parser(PATH):
    out = { "Include RA": {},
            "No RA": {}}
    r = re.compile(r"\((.*?)\)")
    for i in os.listdir(PATH):
        if (("eq 0" in i) or ("lt 0" in i) or ("gt 0" in i)) and ".csv" in i:
            if "Include RA" in i:
                result = r.search(i)
                if result.group(1) in out['Include RA']:
                    print('ERROR: There is a', result.group(1), "in Include RA" )
                    raise Exception
                else:
                    out['Include RA'][result.group(1)] = i
            else:
                result = r.search(i)
                if result.group(1) in out['No RA']:
                    print('ERROR: There is a', result.group(1), "in No RA" )
                    raise Exception
                else:
                    out['No RA'][result.group(1)] = i
    return out

def csv_parser(PATH,dirs):
    frames = []
    for attr,value in dirs.items():
        for filekey, fileloc in dirs[attr].items():
            temp = pd.read_csv(PATH+"/"+fileloc)
            temp["Report Type"] = filekey
            temp['RA Type'] = attr
            temp["StockValue"] = temp['StockValue'].astype(float).fillna(0)
            frames.append(temp)
    result = pd.concat(frames)
    return result

def pivot_table(df):
    out = []
    for ra_type in df['RA Type'].unique():
        for report_type in df['Report Type'].unique():
            temp = df[(df["RA Type"] == ra_type) & (df['Report Type']==report_type)]
            great_zero = temp[temp['StockValue']>0]
            less_zero = temp[temp['StockValue']<0]
            if great_zero.empty == False:
                result = great_zero.groupby(['Category']).sum()
                out.append((report_type+ " "+ ra_type+" g0", result[['StockValue']].round(2)))
            if less_zero.empty == False:
                result = less_zero.groupby(['Category']).sum()
                out.append((report_type + " "+ ra_type+" l0", result[['StockValue']].round(2)))
    return out

def main_table(df):
    tot = df.groupby(["RA Type", "Report Type"]).sum()[["StockQty", "StockValue"]]
    great_zero = df[df["StockValue"]>0].groupby(["RA Type", "Report Type"]).sum()[["StockQty", "StockValue"]]
    less_zero = df[df["StockValue"]<0].groupby(["RA Type", "Report Type"]).sum()[["StockQty", "StockValue"]]
    tot_lab = df[df["Category"] == "LABOUR"].groupby(["RA Type", "Report Type"]).sum()[["StockQty", "StockValue"]]
    tot = tot.rename(columns = {"StockQty": "ITEM QTY", "StockValue": "Value"})
    great_zero = great_zero.rename(columns = {"StockQty": "ITEM QTY", "StockValue": "Value>0"})
    less_zero = less_zero.rename(columns = {"StockQty": "ITEM QTY", "StockValue": "Value<0"})
    tot_lab = tot_lab.rename(columns = {"StockQty": "ITEM QTY", "StockValue": "LA Value"})
    tot = tot.merge(great_zero, on= ["RA Type", "Report Type"], how = "outer")
    tot = tot.merge(less_zero, on= ["RA Type", "Report Type"], how = "outer")
    tot = tot.merge(tot_lab, on= ["RA Type", "Report Type"], how = "outer")
    tot = tot.rename(columns = {"ITEM QTY_x": "ITEM QTY", "ITEM QTY_y": "ITEM QTY"})
    return ("Report", tot.round(2))


def df_to_excel(main, pivot):
    with pd.ExcelWriter("output.xlsx") as writer:
        main[1].to_excel(writer, sheet_name = main[0])    
        for i in pivot:
            i[1].to_excel(writer, sheet_name = i[0])    

    
dirs = dir_parser(PATH)
df = csv_parser(PATH,dirs)
pivot = pivot_table(df)
main = main_table(df)
df_to_excel(main, pivot)

