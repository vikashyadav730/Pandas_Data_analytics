import pandas as pd
import json

file_name = 'Book1.xlsm'
sheet_name = '323. Square Footage'
version = "version"


def main():
    df = pd.read_excel(file_name, sheet_name = sheet_name)
    df = df.drop(["Unnamed: 1"], 1)
    new_header = df.iloc[0]
    df = df[1:]
    df.columns = new_header
    s = pd.DataFrame(df, columns=['Square Footage', 'col', 'P2', 'P12', 'P17'])
    null = s.drop(['col'], axis=1, inplace=True)
    #print(null)
    df = s.apply(pd.to_numeric)

    df = df.set_index('Square Footage')
    df.index = pd.RangeIndex(start=0, stop=2000+1)
    new_df = df.reindex(pd.RangeIndex(start=1776, stop=2000+1 ), fill_value=None)
    new_df = df.interpolate(method='linear', limit_direction='forward')
    new2 = new_df.round(3)
    json_val = new2.to_json(orient='index')
    sheet_name_ar = sheet_name.split()
    customise_json = {}
    rule_name = "Rule " + sheet_name_ar[0]
    rule_number = sheet_name_ar[1] + sheet_name_ar[2]
    customise_json[version] = {rule_name: {rule_number: json.loads(json_val)}}
    print(customise_json)



if __name__ == "__main__":
    main()




