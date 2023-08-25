import pandas as pd
import numpy


class User:
    def __init__(self, id, original, prior, note):
        self.id = id
        self.original = original
        self.prior = prior
        self.note = note


def get_data(file, onlyOriginal):
    xl = pd.ExcelFile(file)
    sheetsList = list(filter(lambda el: el.count("$") == 1, xl.sheet_names))
    usersDict = {}
    placesDict = {}
    for sheet in sheetsList:
        sheetData = xl.parse(sheet)
        usersDict[sheet] = []
        for row in sheetData.itertuples(index=False):
            if row[0] == "№":
                continue
            id = int(row[1]) % 10 ** 5
            original = row[2] == "Да"
            prior = int(row[4])
            note = int(row[5])
            if not (numpy.isnan(row[-1])):
                places = int(row[-1])
                placesDict[sheet] = places
            user = User(id, original, prior, note)
            usersDict[sheet].append(user)
    if onlyOriginal:
        for key in usersDict:
            usersDict[key] = list(filter(lambda user: user.original, usersDict[key]))
    return (usersDict, placesDict)


def sort_data(usersDict, placesDict):
    fin = {key: [] for key in placesDict}
    uid = {}
    while any(len(fin[key]) < placesDict[key] and len(usersDict[key]) > 0 for key in fin):
        for place in placesDict:
            while placesDict[place] > len(fin[place]) and len(usersDict[place]) > 0:
                user = usersDict[place][0]
                fin[place].append(user)
                usersDict[place].pop(0)
                if user.id in uid:
                    uid[user.id] = min(uid[user.id], user.prior)
                else:
                    uid[user.id] = user.prior
        for place in fin:
            fin[place] = list(filter(lambda el: el.prior <= uid[el.id], fin[place]))
    return fin


def make_table(result):
    dfs = {}
    for key in result:
        ids = [el.id for el in result[key]]
        originals = [el.original for el in result[key]]
        priors = [el.prior for el in result[key]]
        notes = [el.note for el in result[key]]
        dfs[key] = pd.DataFrame({"id": ids, "original": originals, "prior": priors, "notes": notes})

    writer = pd.ExcelWriter('./results.xlsx')
    for sheet_name in dfs.keys():
        dfs[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)

    writer.close()


if __name__ == "__main__":
    usersDict, placesDict = get_data('МГУ_риалэ.xlsx', True)
    result = sort_data(usersDict, placesDict)
    make_table(result)