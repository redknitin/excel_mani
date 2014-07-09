__author__ = 'nitinr'

from models.Department import Department

def load_from_sheet(a_sheet):
    AssetColl = []
    for rowidx in range(a_sheet.nrows):
        if a_sheet.cell(rowidx, 1).value == 'Department' or \
                        a_sheet.cell(rowidx, 1).value == 'Department Code' or \
                        a_sheet.cell(rowidx, 1).value == 'nvarchar' or \
                        a_sheet.cell(rowidx, 1).value == '' or \
                        a_sheet.cell(rowidx, 1).value == 15.0:
            continue

        iterAsset = Department()

        iterAsset.department_code = a_sheet.cell(rowidx, 1).value
        iterAsset.description = a_sheet.cell(rowidx, 2).value
        iterAsset.department_type = a_sheet.cell(rowidx, 3).value
        iterAsset.default_store = a_sheet.cell(rowidx, 4).value

        AssetColl.append(iterAsset)
    return AssetColl