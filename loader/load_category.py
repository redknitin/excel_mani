__author__ = 'nitinr'

from models.Category import Category

def load_from_sheet(a_sheet):
    AssetColl = []
    for rowidx in range(a_sheet.nrows):
        if a_sheet.cell(rowidx, 1).value == 'Class & Category' or \
                        a_sheet.cell(rowidx, 1).value == 'Class Code' or \
                        a_sheet.cell(rowidx, 1).value == 'nvarchar' or \
                        a_sheet.cell(rowidx, 1).value == '' or \
                        a_sheet.cell(rowidx, 1).value == 20.0:
            continue

        iterAsset = Category()

        iterAsset.class_code = a_sheet.cell(rowidx, 1).value
        iterAsset.description = a_sheet.cell(rowidx, 2).value
        iterAsset.category_code = a_sheet.cell(rowidx, 3).value
        iterAsset.category_description = a_sheet.cell(rowidx, 4).value

        AssetColl.append(iterAsset)
    return AssetColl