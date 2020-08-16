import pandas as pd 
import numpy as np

gfile = 'Search keyword report.xlsx'
gsheet = 'Sheet0'
bfile = 'keyword-364da045-b456-b8d2-d759-9d63d7dbc8de.xlsx'
bsheet = 'data'
sqlfile = 'report1594866995932.xlsx'
sqlsheet = 'report1594866995932'
finalfile = '719-5.xlsx'
region = ['US']

def goutput():
    dfraw = pd.read_excel(gfile,sheet_name=gsheet,header=2)
    dfraw.columns = dfraw.columns.str.replace(' ','_')

    df = dfraw


    df.dropna(subset=['Keyword'], inplace=True)
    df.replace("Exact match","e",inplace=True)
    df.replace("Phrase match","p",inplace=True)
    df.replace("Broad match","b",inplace=True)
    df.replace("Computers","c",inplace=True)
    df.replace("Tablets","t",inplace=True)
    df.replace("Mobile phones","m",inplace=True)
    df.replace("< 10%","0.05",inplace=True)
    df.replace("> 90%","0.95",inplace=True)
    df.replace("--","0.99123",inplace=True)

    for i in range(len(df)):
    
        df.at[i,"ID"] = str.strip(str.upper(str(df.at[i,"Keyword"])).replace("[",""))
        df.at[i,"ID"] = str.strip(str.upper(str(df.at[i,"ID"])).replace("+",""))
        df.at[i,"ID"] = str.strip(str.upper(str(df.at[i,"ID"])).replace("]",""))
        df.at[i,"ID"] = str.lower(str.upper(str(df.at[i,"ID"])).replace('"',""))


    for i in range(len(df)):
        
        df.at[i,"Keyword"] = str.lower(str(df.at[i,"Keyword"])).replace("[","")
        df.at[i,"Keyword"] = str.lower(str(df.at[i,"Keyword"])).replace("]","")
        df.at[i,"Keyword"] = str.lower(str(df.at[i,"Keyword"])).replace("+","")
        df.at[i,"Keyword"] = str.lower(str(df.at[i,"Keyword"])).replace(" ","")
        df.at[i,"Keyword"] = str.lower(str(df.at[i,"Keyword"])).replace('"',"")
    
        


    dfm = df[df.Device.isin(["m"])]
    dfc = df[df.Device.isin(["c"])]
    dft = df[df.Device.isin(["t"])]

    dfnew = dfm.append(dfc.append(dft))

    dtypedic = {'Impr._(Top)_%':float,
                'Click_share':float,
                'Search_impr._share':float,
                'Search_top_IS':float,
                'Search_exact_match_IS':float,
                'Search_lost_top_IS_(rank)':float,
        }



    dfnew = dfnew.astype(dtypedic)
    dfnew['Vendor'] = 'Google'

    dfnew['PI'] = dfnew['Impr.']/dfnew['Search_impr._share']
    dfnew['PTI'] = dfnew['Impr.']/dfnew['Search_top_IS']
    dfnew['adgroup'] = dfnew['Match_type'] + '--' + dfnew['Device'] + '-' + dfnew['ID'] + '-' + dfnew['Vendor']


    AWresult = dfnew.groupby(['adgroup','ID']).agg(
    {
        'Impr.':sum,
        'Clicks':'sum',
        'Conversions':'sum',
        'PI':'sum',
        'PTI':'sum',
        'Cost':'sum'
        
    })
    AWresult=AWresult.reset_index()
    return AWresult








def boutput():
    dfraw = pd.read_excel(bfile,sheet_name=bsheet,header=3)
    bingcols = ['Device','Keyword_status','Keyword', 'Campaign', 'Ad_group','Match_type','Max._CPC','Cost', 'Clicks', 'Conversions','Impr.','Search_impr._share','Click_share','Search_top_IS','Search_lost_top_IS_(rank)','Search_exact_match_IS']
    dfraw.columns = bingcols

    df = dfraw


    df.dropna(subset=['Keyword'], inplace=True)
    df.replace("Exact","e",inplace=True)
    df.replace("Phrase","p",inplace=True)
    df.replace("Broad","b",inplace=True)
    df.replace("Computer","c",inplace=True)
    df.replace("Tablet","t",inplace=True)
    df.replace("SmartPhone","m",inplace=True)
    df.replace("0%","0.0001",inplace=True)
    df.replace("> 90%","0.95",inplace=True)
    df.replace("-","0.99123",inplace=True)

    for i in range(len(df)):
    
        df.at[i,"ID"] = str.strip(str.upper(str(df.at[i,"Keyword"])).replace("[",""))
        df.at[i,"ID"] = str.strip(str.upper(str(df.at[i,"ID"])).replace("+",""))
        df.at[i,"ID"] = str.strip(str.upper(str(df.at[i,"ID"])).replace("]",""))
        df.at[i,"ID"] = str.lower(str.upper(str(df.at[i,"ID"])).replace('"',""))


    for i in range(len(df)):
        
        df.at[i,"Keyword"] = str.lower(str(df.at[i,"Keyword"])).replace("[","")
        df.at[i,"Keyword"] = str.lower(str(df.at[i,"Keyword"])).replace("]","")
        df.at[i,"Keyword"] = str.lower(str(df.at[i,"Keyword"])).replace("+","")
        df.at[i,"Keyword"] = str.lower(str(df.at[i,"Keyword"])).replace(" ","")
        df.at[i,"Keyword"] = str.lower(str(df.at[i,"Keyword"])).replace('"',"")
    
        


    dfm = df[df.Device.isin(["m"])]
    dfc = df[df.Device.isin(["c"])]
    dft = df[df.Device.isin(["t"])]

    dfnew = dfm.append(dfc.append(dft))

    dfnew = dfm.append(dfc)

    dtypedic = {
                'Click_share':float,
                'Search_impr._share':float,
                'Search_top_IS':float,
                'Search_exact_match_IS':float,
                'Search_lost_top_IS_(rank)':float,
        }



    dfnew = dfnew.astype(dtypedic)

    dfnew['Vendor'] = 'Bing'

    dfnew['PI'] = dfnew['Impr.']/dfnew['Search_impr._share']
    dfnew['PTI'] = dfnew['Impr.']/dfnew['Search_top_IS']
    dfnew['adgroup'] = dfnew['Match_type'] + '--' + dfnew['Device'] + '-' + dfnew['ID'] + '-' + dfnew['Vendor']


    result2 = dfnew.groupby(['adgroup','ID']).sum()

    result2 = result2.reset_index()

    result2['Vendor'] = 'Bing'






    Bresult = result2.loc[:,['adgroup', 'Impr.', 'Clicks', 'Conversions', 'PI', 'PTI', 'Cost','ID']]
    return Bresult

def SQLoutput():
    import pandas as pd 
    import numpy as np

    def sizecap (i):
        if i>100:
            return 100
        else:
            return i

    
    df_file_1 = pd.read_excel(sqlfile,sheet_name=sqlsheet)



    for i in range(len(df_file_1)):

        df_file_1.at[i,"Media Keyword"] = str.lower(str(df_file_1.at[i,"Media Keyword"])).replace("[","")
        df_file_1.at[i,"Media Keyword"] = str.lower(str(df_file_1.at[i,"Media Keyword"])).replace("]","")
        df_file_1.at[i,"Media Keyword"] = str.lower(str(df_file_1.at[i,"Media Keyword"])).replace('+',"")
        df_file_1.at[i,"Media Keyword"] = str.strip(str(df_file_1.at[i,"Media Keyword"])).replace('"',"")

    c2= df_file_1.loc[:,['Media Keyword Match Type','Media Keyword','Devices','Potential /Addressable Fleet Size','Vendor','Converted','Region']]


    c2['adgroup'] = c2['Media Keyword Match Type'] + '--' + c2['Devices'] + '-' + c2['Media Keyword'] + '-' + c2['Vendor']  
    c2['FS'] = c2['Potential /Addressable Fleet Size'].apply(sizecap)

    #take out non-SQL fleet size
    # for i in range(len(c2)):
    #     if c2.at[i,"Converted"] == 0:
    #         c2.at[i,"FS"] = 0
    #     else:
    #         c2.at[i,"FS"] = c2.at[i,"FS"]
        


    #Country Filter
    c3 = c2[c2.Region.isin(region)]

    SQLresult = c3.groupby(['adgroup']).agg(
    {'Converted':sum,
        'Vendor':'count',
        'FS':'sum'
        
    })
    SQLresult = SQLresult.reset_index()
    return SQLresult

def main():
    AWresult = goutput()
    Bresult = boutput()
    SQLresult = SQLoutput()
   

    
    AWresult = AWresult.append(Bresult)

    big = pd.merge(SQLresult,AWresult,how='outer')
    big['PI'] = pd.Series([round(val, 0) for val in big['PI']], index = big.index)
    big['PTI'] = pd.Series([round(val, 0) for val in big['PTI']], index = big.index)
    big['Cost'] = pd.Series([round(val, 0) for val in big['Cost']], index = big.index)
    big['CSQL'] = big['Cost']/big['Converted']
    big['CSQL'] = pd.Series([round(val, 0) for val in big['CSQL']], index = big.index)
    big['CPC'] = big['Cost']/big['Clicks']
    big['CPC'] = pd.Series([round(val, 1) for val in big['CPC']], index = big.index)
    big['CPCon'] = big['Cost']/big['Conversions']
    big['CPCon'] = pd.Series([round(val, 0) for val in big['CPCon']], index = big.index)
    big['AVGFS'] = big['FS']/big['Converted']
    big['AVGFS'] = pd.Series([round(val, 0) for val in big['AVGFS']], index = big.index)


    big['PI%'] = big['Impr.']/big['PI']
    
    big.head
    big['PTI%'] = big['Impr.']/big['PTI']
    

    big['SQL%'] = big['Converted']/big['Clicks']
    

    big['Qual%'] = big['Converted']/big['Vendor']
    

    big['Raw%'] = big['Vendor']/big['Conversions']
    

    colrenamedic = {'Converted':'SQL', 'Vendor':'Raw', 'Conversions':'Conv.'}
    big.rename(columns = colrenamedic,inplace=True)

    
    
    big.to_excel(finalfile)

main()

