import pandas as pd
import openpyxl
import warnings
import re

def data_frame_from_xlsx(xlsx_file, range_name,header_row=False, shorten_names=True, index_col=False):
    """ Get a single rectangular region from the specified file.
    range_name can be a standard Excel reference ('Sheet1!A2:B7') or 
    refer to a named region ('my_cells')."""
    wb = openpyxl.load_workbook(xlsx_file, data_only=True, read_only=True)
    if '!' in range_name:
        # passed a worksheet!cell reference
        ws_name, reg = range_name.split('!')
        if ws_name.startswith("'") and ws_name.endswith("'"):
            # optionally strip single quotes around sheet name
            ws_name = ws_name[1:-1]
        region = wb[ws_name][reg]
    else:
        # passed a named range; find the cells in the workbook
        full_range = wb.defined_names[range_name]
        if full_range is None:
            raise ValueError(
                'Range "{}" not found in workbook "{}".'.format(range_name, xlsx_file)
            )
        # convert to list (openpyxl 2.3 returns a list but 2.4+ returns a generator)
        destinations = list(full_range.destinations) 
        if len(destinations) > 1:
            raise ValueError(
                'Range "{}" in workbook "{}" contains more than one region.'
                .format(range_name, xlsx_file)
            )
        ws, reg = destinations[0]
        # convert to worksheet object (openpyxl 2.3 returns a worksheet object 
        # but 2.4+ returns the name of a worksheet)
        if isinstance(ws, str):
            ws = wb[ws]
        region = ws[reg]
    df = pd.DataFrame([cell.value for cell in row] for row in region)
    
    if header_row:
        header = df.iloc[0]
        df.columns = header
        # Strip off everything after the first punctuation
        if shorten_names:
            df.columns = header.apply(lambda x : re.sub(r'[)( /\*].*$', '', x) if x else None)

        df=df.drop(0)
        df.index = range(len(df))

    if index_col:
        df.set_index([index_col],drop=False,inplace=True)

    return df
