# Python program to read json file - party list as example, and export into excel

# import library for working with json
import json

# import pandas for working with data and export to excel file
import pandas as pd
 
# Opening JSON file - for reading - as example is used party list
f = open("JsonParserPython\\PartyList.json")
 
# returns JSON object as a dictionary
data = json.load(f)
print("Dictionary created from json...")

# choose "my_party_list"
party_list = data["my_party_list"]
 
# Iterating through the json list
for i in data['my_party_list']:
    print(i["first_name"] + "\t" + i["last_name"] + "\t" + i["phone"] + "\t" + i["email"])
    # rename keys in dictionary
    i["First Name"] = i["first_name"]
    del i["first_name"]
    i["Last Name"] = i["last_name"]
    del i["last_name"]
    i["Phone"] = i["phone"]
    del i["phone"]
    i["Email"] = i["email"]
    del i["email"]


# Creating Excel Writer Object from Pandas  
writer = pd.ExcelWriter('JsonParserPython\\party_list.xlsx',engine='xlsxwriter')   

df = pd.DataFrame.from_dict(party_list)
df.to_excel(writer,sheet_name='Party list',startrow=0 , startcol=0)

# Closing excel file
writer.close()

print("Dictionary converted into excel...")

# Closing json file
f.close()