import pandas as pd 
import numpy as np



file = 'Search keyword report.csv'
dfraw = pd.read_csv(file)
dfraw.columns = dfraw.columns.str.replace(' ','_')

# df = dfraw


# df.dropna(subset=['Keyword'], inplace=True)
# df.replace("Exact match","e",inplace=True)
# df.replace("Phrase match","p",inplace=True)
# df.replace("Broad match","b",inplace=True)
# df.replace("Computers","c",inplace=True)
# df.replace("Tablets","m",inplace=True)
# df.replace("Mobile phones","m",inplace=True)
# df.replace("< 10%","0.05",inplace=True)
# df.replace("> 90%","0.95",inplace=True)
# df.replace("--","0.99123",inplace=True)


# for i in range(len(df)):
    
#     df.at[i,"Keyword"] = str.lower(str(df.at[i,"Keyword"])).replace("[","")
#     df.at[i,"Keyword"] = str.lower(str(df.at[i,"Keyword"])).replace("]","")
#     df.at[i,"Keyword"] = str.lower(str(df.at[i,"Keyword"])).replace("+","")
#     df.at[i,"Keyword"] = str.lower(str(df.at[i,"Keyword"])).replace(" ","")
#     df.at[i,"Keyword"] = str.lower(str(df.at[i,"Keyword"])).replace('"',"")
  
    


# dfm = df[df.Device.isin(["m"])]
# dfc = df[df.Device.isin(["c"])]

# dfnew = dfm.append(dfc)

# dtypedic = {'Impr._(Top)_%':float,
#             'Click_share':float,
#             'Search_impr._share':float,
#             'Search_top_IS':float,
#             'Search_exact_match_IS':float,
#             'Search_lost_top_IS_(rank)':float,
#     }



# dfnew = dfnew.astype(dtypedic)
# dfnew['Vendor'] = 'Google'

# dfnew['PI'] = dfnew['Impr.']/dfnew['Search_impr._share']
# dfnew['PTI'] = dfnew['Impr.']/dfnew['Search_top_IS']
# dfnew['ID'] = dfnew['Match_type'] + '-' + dfnew['Device'] + '-' + dfnew['Keyword'] + '-' + dfnew['Vendor']


# AWresult = dfnew.groupby('ID').agg(
# {
#     'Impr.':sum,
#     'Clicks':'sum',
#     'Conversions':'sum',
#     'PI':'sum',
#     'PTI':'sum',
#     'Cost':'sum'
    
# })

# AWresult.to_excel("189.xlsx")