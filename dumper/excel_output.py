__author__ = 'nitinr'

from xlwt3 import Workbook


def write_assets(a_sheet, a_coll):
    counter = 0
    colTitleArr = ['Asset Code', 'FSI Asset Code', 'Bar Code', 'Description', 'Department', 'Class', 'Category', 'Parent', 'Location', 'Install Date', 'Cost Code', 'Criticality', 'Manufacturer', 'Serial No', 'Purchase Value', 'Assigned To', 'Vendor']
    for colHeadings in colTitleArr:
        a_sheet.write(0, counter, colHeadings)
        counter += 1

    counter = 1
    for iterrow in a_coll:
        for iterx in range(len(colTitleArr)):
            rowdata = [
                iterrow.asset_code,
                iterrow.fsi_asset_code,
                iterrow.bar_code,
                iterrow.description,
                iterrow.department,
                iterrow.class_code,
                iterrow.category_code,
                iterrow.parent_asset,
                iterrow.location_code,
                iterrow.asset_install_date,
                iterrow.cost_code,
                iterrow.asset_criticality,
                iterrow.manufacturer_code,
                iterrow.serial_no,
                iterrow.purchase_price,
                iterrow.assigned_to,
                iterrow.vendor_supplier
            ]
            a_sheet.write(counter, iterx, rowdata[iterx])
        counter += 1


def write_prop(a_sheet, a_coll):
    counter = 0
    colTitleArr = ['Asset Code', 'Description', 'Department', 'Class', 'Category', 'Parent', 'Location', 'Install Date', 'Cost Code', 'Purchase Value', 'Assigned To']
    for colHeadings in colTitleArr:
        a_sheet.write(0, counter, colHeadings)
        counter += 1

    counter = 1
    for iterrow in a_coll:
        for iterx in range(len(colTitleArr)):
            rowdata = [
                iterrow.asset_code,
                iterrow.description,
                iterrow.department,
                iterrow.class_code,
                iterrow.category_code,
                iterrow.parent_asset,
                iterrow.location_code,
                iterrow.asset_install_date,
                iterrow.cost_code,
                iterrow.purchase_price,
                iterrow.assigned_to
            ]
            a_sheet.write(counter, iterx, rowdata[iterx])
        counter += 1


def write_dept(a_sheet, a_coll):
    counter = 0
    colTitleArr = ['Department Code', 'Description', 'Type', 'Default Store']
    for colHeadings in colTitleArr:
        a_sheet.write(0, counter, colHeadings)
        counter += 1

    counter = 1
    for iterrow in a_coll:
        for iterx in range(len(colTitleArr)):
            rowdata = [
                iterrow.department_code,
                iterrow.description,
                iterrow.department_type,
                iterrow.default_store,
            ]
            a_sheet.write(counter, iterx, rowdata[iterx])
        counter += 1


def write_cat(a_sheet, a_coll):
    counter = 0
    colTitleArr = ['Class Code', 'Description', 'Category', 'Category Description']
    for colHeadings in colTitleArr:
        a_sheet.write(0, counter, colHeadings)
        counter += 1

    counter = 1
    for iterrow in a_coll:
        for iterx in range(len(colTitleArr)):
            rowdata = [
                iterrow.class_code,
                iterrow.description,
                iterrow.category_code,
                iterrow.category_description,
            ]
            a_sheet.write(counter, iterx, rowdata[iterx])
        counter += 1


def write_to_file(a_filename, asset_coll, prop_coll, cat_coll, dept_coll):
    wb = Workbook()

    asset_sheet = wb.add_sheet('Assets')
    write_assets(asset_sheet, asset_coll)

    prop_sheet = wb.add_sheet('Property Master')
    write_prop(prop_sheet, prop_coll)

    cat_sheet = wb.add_sheet('Category & Class')
    write_cat(cat_sheet, cat_coll)

    dept_sheet = wb.add_sheet('Department')
    write_dept(dept_sheet, dept_coll)

    wb.save(a_filename)