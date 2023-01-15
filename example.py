from openpyxl_style_writer import RowWriter
from openpyxl_style_writer import DefaultStyle, CustomStyle


if __name__ == '__main__':
    workbook = RowWriter()
    workbook.create_sheet('ExampleSheet')

    title = 'This is an example'
    row_title_1 = ['fruits', 'fruits', 'animals', 'animals']
    row_title_2 = ['apple', 'banana', 'cat', 'dog']
    data = [10, 20, 30, 40]

    # append single cell with Default Style
    workbook.row_append(title)
    workbook.set_cell_width(1, 30)
    workbook.create_row()
    for item in row_title_1:
        workbook.row_append(item)
    workbook.create_row()

    # set custom Default Style and append list in a row
    blue_font_style = {'color': '0000ff', 'bold': True, 'size': 8}
    DefaultStyle.set_default(font_params=blue_font_style)
    workbook.row_append_list(row_title_2)
    workbook.create_row()

    # create new Custom Style and give row_append_list a style
    pink_fill_style = {'patternType': 'solid', 'fgColor': 'd25096'}
    pink_style = CustomStyle(fill_params=pink_fill_style)
    workbook.row_append_list(data, pink_style)
    workbook.create_row()

    workbook.save('example.xlsx')
