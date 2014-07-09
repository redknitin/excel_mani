__author__ = 'nitinr'


from configparser import ConfigParser
from xlrd import open_workbook

from loader import load_assets
from loader import load_property
from loader import load_dept
from loader import load_category
from dumper import excel_output
from dumper import csv_output

# from multiprocessing.pool import ThreadPool

cfg = ConfigParser()
cfg.read('settings.ini')


def lookup_asset(a_code, asset_col):
    for iter_row in asset_col:
        if iter_row.asset_code == a_code:
            return iter_row
    return None


def split_till_found(a_code, a_col):
    parts = a_code.split('-')
    if len(parts) == 1: return ''
    par_code = '-'.join(parts[:-1])
    if lookup_asset(par_code, a_col):
        return par_code
    else:
        return split_till_found(par_code, a_col)


def process_for_prop_parent(a_col):
    for iter_row in a_col:
        if iter_row.parent_asset == '':
            iter_row.parent_asset = split_till_found(iter_row.asset_code, a_col)
        if iter_row.parent_asset == '': print('Couldnt find parent for '+iter_row.asset_code)
        pass


def process_for_asset_dup(a_col):
    for iteridx in range(len(a_col)):
        if iteridx > len(a_col)-1:
            break
        iterobj = a_col[iteridx]
        lst = list(filter(lambda x: x.asset_code == iterobj.asset_code, a_col))
        if len(lst) > 1:
            a_col.remove(iterobj)


def process_for_cat_dup(a_col):
    for iteridx in range(len(a_col)):
        if iteridx > len(a_col)-1:
            break
        iterobj = a_col[iteridx]
        lst = list(filter(lambda x: x.category_code == iterobj.category_code, a_col))
        if len(lst) > 1:
            a_col.remove(iterobj)


def process_for_dept_dup(a_col):
    for iteridx in range(len(a_col)):
        if iteridx > len(a_col)-1:
            break
        iterobj = a_col[iteridx]
        lst = list(filter(lambda x: x.department_code == iterobj.department_code, a_col))
        if len(lst) > 1:
            a_col.remove(iterobj)


input_filename = cfg['Input']['Filename']

asset_col = []
dept_col = []
cat_col = []
prop_col = []

for iter_filename in input_filename.split(';'):
    wb = open_workbook(iter_filename)

    s_asset = wb.sheet_by_name(cfg['Input']['AssetSheetname'])
    asset_coll = load_assets.load_from_sheet(s_asset)
    asset_col += asset_coll

    s_prop = wb.sheet_by_name(cfg['Input']['PropertySheetname'])
    prop_coll = load_property.load_from_sheet(s_prop)
    prop_col += prop_coll

    s_dept = wb.sheet_by_name(cfg['Input']['DepartmentSheetname'])
    dept_coll = load_dept.load_from_sheet(s_dept)
    dept_col += dept_coll

    s_cat = wb.sheet_by_name(cfg['Input']['CategorySheetname'])
    cat_coll = load_category.load_from_sheet(s_cat)
    cat_col += cat_coll


process_for_asset_dup(asset_col)
process_for_asset_dup(prop_col)
process_for_cat_dup(cat_col)
process_for_dept_dup(dept_col)

process_for_prop_parent(prop_col)

excel_output.write_to_file(cfg['Output']['Filename'], asset_col, prop_col, cat_col, dept_col)




#csv_output.write_to_file(cfg['Output']['Filename'], asset_col, prop_col, cat_col, dept_col)