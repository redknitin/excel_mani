__author__ = 'nitinr'

from models.Property import Property

def load_from_sheet(a_sheet):
    AssetColl = []
    for rowidx in range(a_sheet.nrows):
        if a_sheet.cell(rowidx, 0).value == 'HOME' or \
                        a_sheet.cell(rowidx, 0).value == 'Asset Code' or \
                        a_sheet.cell(rowidx, 0).value == 'nvarchar' or \
                        a_sheet.cell(rowidx, 0).value == '' or \
                        a_sheet.cell(rowidx, 0).value == 30.0:
            continue

        iterAsset = Property()

        iterAsset.asset_code = a_sheet.cell(rowidx, 0).value
        iterAsset.description = a_sheet.cell(rowidx, 1).value
        iterAsset.department = a_sheet.cell(rowidx, 2).value
        iterAsset.class_code = a_sheet.cell(rowidx, 3).value
        iterAsset.category_code = a_sheet.cell(rowidx, 4).value
        iterAsset.parent_asset = a_sheet.cell(rowidx, 5).value
        iterAsset.location_code = a_sheet.cell(rowidx, 6).value
        iterAsset.asset_install_date = a_sheet.cell(rowidx, 7).value
        iterAsset.cost_code = a_sheet.cell(rowidx, 8).value
        iterAsset.purchase_price = a_sheet.cell(rowidx, 9).value
        iterAsset.assigned_to = a_sheet.cell(rowidx, 10).value

        AssetColl.append(iterAsset)
    return AssetColl