import codecs
import os
from os import walk
from xml.etree import ElementTree
import pandas as pd
from pathlib import Path

ENCODING = "cp932"
SEP = ","

DIR_LIST = "C:/Users/digiworker_biz_02/Desktop/mfz_input/"
# APP_TYPE = "交際費(仮払)申請書"
APP_TYPE = '経費(仮払)申請書'
# APP_TYPE = '支払依頼書'
# Recursively find the xml files
for file_path in Path(DIR_LIST).glob("**/*.xml"):

    split_path = os.path.splitext(file_path)
    file_header = split_path[0]
    tree = ElementTree.parse(file_path)

    root = tree.getroot()
    main_array = []
    child_header_array = []
    child_value_array = []
    for child in root:
        my_dict = child.attrib
        header_name = my_dict["id"]
        child_header_array.append(header_name)
        child_value_array.append(child.text)

    main_array.append(child_header_array)
    main_array.append(child_value_array)

    df = pd.DataFrame(main_array)
    # Create Dir
    os.makedirs(file_header, exist_ok=True)
    # needed this trick to get file name
    file_path_base = os.path.basename(file_header)

    # Explode with ~ character
    df = df.astype(str)
    #
    # col = df.pop("DebtKindInfo")
    # df.insert(0, col.name, col)
    #
    # for itm in df.head():
    #     df[itm] = df[itm].str.split('~')
    # df = df.apply(pd.Series.explode)

    df.to_csv(
        file_header + "/" + file_path_base + ".csv",
        sep=SEP,
        encoding=ENCODING,
        index=False,
        header=False,
    )


def open_csv(csv_file):
    with codecs.open(csv_file, "r", ENCODING, "ignore") as file:
        return pd.read_csv(file, sep=SEP)


# Further more can combine multiple csvs

dfs = (open_csv(p) for p in Path(DIR_LIST).glob("**/" + APP_TYPE + "*.csv"))
res = pd.concat(dfs)

df = res

col = df.pop('DebtKindInfo')
df.insert(0, col.name, col)

print(df)
df['DebtKindInfo'] = df['DebtKindInfo'].str.split('~')
print(df)

df = df.apply(pd.Series.explode)

# export to csv
merged_csv = DIR_LIST + "/MERG_" + APP_TYPE + ".csv"
df.to_csv(merged_csv, sep=SEP, index=False, encoding=ENCODING)
