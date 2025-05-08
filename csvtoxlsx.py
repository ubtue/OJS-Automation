import pandas as pd
import datetime
import ast

date = datetime.datetime.today().strftime("%Y%m%d")

filename = f"serverFile"

sheetName = "Info"

content = pd.read_csv(f"{filename}.csv", sep=";")

writer = pd.ExcelWriter(f"{filename}.xlsx", mode="w")

content.to_excel(writer, merge_cells=False, index=False, sheet_name=sheetName)

sheet = writer.sheets["Info"]

columnLength = 25

urlList = []
expiryDays = []
vm = []

for index, entry in enumerate(content["URLs"]):
    for url in ast.literal_eval(entry):
        urlList.append(url)
        vm.append(content["Hostname"][index].split(".")[0])

for entry in content["expiration days"]:
    for days in ast.literal_eval(entry):
        expiryDays.append(days)


certificates = pd.DataFrame(
    {"URLs": urlList, "Expiration Days": expiryDays, "Machine": vm}
)

certificates.to_excel(
    writer,
    merge_cells=False,
    index=False,
    sheet_name=sheetName,
    startrow=len(content) + 2,
)

for column in content:
    columnIdx = content.columns.get_loc(column)
    sheet.set_column(columnIdx, columnIdx, columnLength)

writer.close()
