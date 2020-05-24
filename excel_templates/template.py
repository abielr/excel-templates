from copy import copy
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter

class ExcelTemplate(object):
    def __init__(self, filename):
        self.wb = load_workbook(filename)
        self.wbv = load_workbook(filename, data_only=True)
        self.original_dimensions = dict((ws.title, (ws.max_row, ws.max_column)) for ws in self.wb.worksheets)
        self.updated_dimensions = {}
        self.cell_dimensions = dict((ws.title, (1,1)) for ws in self.wb.worksheets)

    def copy_worksheet(self, source_name, target_name):
        target = self.wb.copy_worksheet(self.wb[source_name])
        target.title = target_name

        target = self.wbv.copy_worksheet(self.wbv[source_name])
        target.title = target_name

        self.original_dimensions[target_name] = self.original_dimensions[source_name]
        self.updated_dimensions[target_name] = self.updated_dimensions.get(source_name)
        self.cell_dimensions[target_name] = self.cell_dimensions[source_name]

    def tile(self, sheetname, rows, columns, row_spacing=1, col_spacing=1):
        if sheetname in self.updated_dimensions:
            raise Exception("Can only expand the sheet '{sheet}' once".format(sheet=sheetname))

        ws = self.wb[sheetname]
        rng = ws[ws.dimensions]

        row_offset = ws.max_row + row_spacing
        col_offset = ws.max_column + col_spacing
        self.updated_dimensions[sheetname] = (row_offset, col_offset)
        self.cell_dimensions[sheetname] = (rows, columns)

        merged_cells = [(r.min_row, r.min_col, r.max_row, r.max_col) for r in ws.merged_cells.ranges]

        for j in range(columns):
            for i in range(rows):
                if i == 0 and j == 0: continue
                for row in rng:
                    for source_cell in row:
                        target_cell = ws.cell(source_cell.row + row_offset*i, source_cell.column + col_offset*j)

                        if not isinstance(source_cell.value, int):
                            target_cell._value = Translator(source_cell._value, origin=source_cell.coordinate). \
                                translate_formula(target_cell.coordinate)
                        else:
                            target_cell.value = source_cell.value
                        target_cell.data_type = source_cell.data_type

                        if source_cell.has_style:
                            target_cell._style = copy(source_cell._style)

                        if source_cell.hyperlink:
                            target_cell._hyperlink = copy(source_cell.hyperlink)

                        if source_cell.comment:
                            target_cell.comment = copy(source_cell.comment)

                for min_row, min_col, max_row, max_col in merged_cells:
                    ws.merge_cells(start_row=min_row+row_offset*i, start_column=min_col+col_offset*j,
                                   end_row=max_row+row_offset*i, end_column=max_col+col_offset*j)

        for i in range(1,rows):
            for ii in range(len(rng)):
                ws.row_dimensions[ii+1+row_offset*i].height = ws.row_dimensions[ii+1].height

        for j in range(1,columns):
            for jj in range(len(rng[0])):
                ws.column_dimensions[get_column_letter(jj+col_offset*j+1)].width = \
                    ws.column_dimensions[get_column_letter(jj + 1)].width

    def fill(self, sheetname, data, grid_row=1, grid_col=1, prefix='', fillna=None):
        if grid_row > self.cell_dimensions[sheetname][0] or grid_col > self.cell_dimensions[sheetname][1]:
            raise Exception("Invalid cell position (%d, %d)" % (grid_row, grid_col))
        ws = self.wb[sheetname]
        row_offset, col_offset = self.updated_dimensions.get(sheetname, self.original_dimensions[sheetname])
        orow_offset, ocol_offset = self.original_dimensions[sheetname]

        for irow in range(1, orow_offset+1):
            for icol in range(1, ocol_offset+1):
                newrow = irow + (grid_row-1)*row_offset
                newcol = icol + (grid_col-1)*col_offset
                key = str(self.wbv[sheetname].cell(irow, icol).value)
                if key.startswith(prefix):
                    if key[len(prefix):] in data:
                        ws.cell(newrow, newcol).value = data[key[len(prefix):]]
                    elif not fillna is None:
                        ws.cell(newrow, newcol).value = fillna

    def save(self, filename):
        self.wb.save(filename)

def make_dict(df, keys, value, sep):
    """
    Turn a DataFrame into a dictionary by joining column values
    :param df: input DataFrame
    :param keys: list of column names to join together
    :param value: column containing values
    :param sep: string to join on
    :param prefix: a prefix to append to the start of every key
    :return: a dictionary
    """
    d = {}
    for i, row in df.iterrows():
        d[sep.join([str(row[key]) for key in keys])] = row[value]
    return d
