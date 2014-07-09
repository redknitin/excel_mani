__author__ = 'nitinr'

from models.Asset import Asset

def load_from_sheet(a_sheet):
    AssetColl = []
    for rowidx in range(a_sheet.nrows):
        if a_sheet.cell(rowidx, 0).value == 'HOME' or \
                        a_sheet.cell(rowidx, 0).value == 'Asset Code' or \
                        a_sheet.cell(rowidx, 0).value == 'nvarchar' or \
                        a_sheet.cell(rowidx, 0).value == '' or \
                        a_sheet.cell(rowidx, 0).value == 30.0:
            continue

        iterAsset = Asset()

        iterAsset.asset_code = a_sheet.cell(rowidx, 0).value
        iterAsset.fsi_asset_code = a_sheet.cell(rowidx, 1).value
        iterAsset.bar_code = a_sheet.cell(rowidx, 2).value
        iterAsset.description = a_sheet.cell(rowidx, 3).value
        iterAsset.department = a_sheet.cell(rowidx, 4).value
        iterAsset.class_code = a_sheet.cell(rowidx, 5).value
        iterAsset.category_code = a_sheet.cell(rowidx, 6).value
        iterAsset.parent_asset = a_sheet.cell(rowidx, 7).value
        iterAsset.location_code = a_sheet.cell(rowidx, 8).value
        iterAsset.asset_install_date = a_sheet.cell(rowidx, 9).value
        iterAsset.cost_code = a_sheet.cell(rowidx, 10).value
        iterAsset.asset_criticality = a_sheet.cell(rowidx, 11).value
        iterAsset.manufacturer_code = a_sheet.cell(rowidx, 12).value
        iterAsset.serial_no = a_sheet.cell(rowidx, 13).value
        iterAsset.purchase_price = a_sheet.cell(rowidx, 14).value
        iterAsset.assigned_to = a_sheet.cell(rowidx, 15).value
        iterAsset.vendor_supplier = a_sheet.cell(rowidx, 16).value

        AssetColl.append(iterAsset)
    return AssetColl