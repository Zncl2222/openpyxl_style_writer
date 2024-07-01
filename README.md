# openpyxl_style_writer

![licence](https://img.shields.io/github/license/Zncl2222/openpyxl_style_writer)
[![ci](https://img.shields.io/github/actions/workflow/status/Zncl2222/openpyxl_style_writer/github-pre-commit.yml?logo=pre-commit&label=pre-commit)](https://github.com/Zncl2222/openpyxl_style_writer/actions/workflows/github-pre-commit.yml)
[![GitHub Workflow Status (with event)](https://img.shields.io/github/actions/workflow/status/Zncl2222/openpyxl_style_writer/pytest.yml?logo=pytest&label=pytest)](https://github.com/Zncl2222/openpyxl_style_writer/actions/workflows/pytest.yml)
[![codecov](https://codecov.io/gh/Zncl2222/openpyxl_style_writer/graph/badge.svg?token=UP6M8SCC14)](https://codecov.io/gh/Zncl2222/openpyxl_style_writer)
![language](https://img.shields.io/badge/Solutions-black.svg?style=flat&logo=python)

This is a wrapper base on [openpyxl](https://pypi.org/project/openpyxl/) package. The original feature to create resuable style ([NameStyled](https://openpyxl.readthedocs.io/en/stable/styles.html#creating-a-named-style)) is not avaliable for [write only mode](https://openpyxl.readthedocs.io/en/stable/optimized.html#write-only-mode). Thus this package aimed to provide a easy way for user to create resuable styles and use it on [write only mode](https://openpyxl.readthedocs.io/en/stable/optimized.html#write-only-mode) easily

# Installation

```$ pip install openpyxl_style_writer ```

# Usage

### Example

```python
from openpyxl_style_writer import CustomStyle, DefaultStyle, RowWriter


if __name__ == '__main__':
    workbook = RowWriter()
    # enable protection by protection=True
    workbook.create_sheet('ExampleSheet', protection=True)

    title = 'This is an example'
    row_title_1 = ['fruits', 'fruits', 'animals', 'animals']
    row_title_2 = ['apple', 'banana', 'cat', 'dog']
    percent_data = [0.1, 0.6, 0.225, 0.4755, 0.9, 1]
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
    # add protect to pink_style
    pink_style = CustomStyle(fill_params=pink_fill_style, protect=True)
    workbook.row_append_list(data, pink_style)
    workbook.create_row()

    # create number_format style
    percent_style = CustomStyle(font_size=8, number_format='0.0%')
    workbook.row_append_list(percent_data, percent_style)
    workbook.create_row()

    workbook.save('example.xlsx')
```

### CustomStyle & DefaultStyle
You can either set the `CustomStyle` and `DefaultStyle` by the key words from openpyxl or from this package. Here is an example to create a `CustomStyle` by the openpyxl key words.

```python
from openpyxl_style_writer import CustomStyle

# Create the style by the original key words from openpyxl

blue_title_font = {
    'color': '0000ff',
    'bold': True,
    'size': 15,
}
cyan_title_pattern = {
    'patternType': 'solid',
    'fgColor': '00ffff	'
}

# Use font_params and fill_params to create a reusable style with the blue font and cyan fill.
# The style you do not set will use the style of DefaultStyle.

custom_title_style = CustomStyle(
    font_params=blue_title_font,
    fill_params=cyan_titl_patter,
)
```
You could also set your `DefaultStyle` first, and the rest of the `CustomStyle` will follow the settings of `DefaultStyle` if the style do not set in `CustomStyle`.

```python
from openpyxl_style_writer import CustomStyle, DefaultStyle

cyan_title_pattern = {
    'patternType': 'solid',
    'fgColor': '00ffff	'
}

blue_title_font = {
    'color': '0000ff',
    'bold': True,
    'size': 15,
}

# set default style with cyan fill
DefaultStyle.set_default(fill_params=cyan_title_pattern)

# This custom style will show blue font and cyan fill, although it only set the font_params
custom_title_style = CustomStyle(
    font_params=blue_title_font,
)
```

If you want to do it in simple way, openpyxl_style_writer offer a map for some [common key words](#list-of-key-words-in-openpyxl_style_writer). You can use the key words of openpyxl_style_writer.
```python
from openpyxl_style_writer import CustomStyle

custom_title_style = CustomStyle(
    font_size=15,
    font_name='Calibri',
)
```

Or you can use both methods
```python
from openpyxl_style_writer import CustomStyle

blue_title_font = {
    'color': '0000ff',
    'bold': True,
    'size': 15,
}

custom_title_style = CustomStyle(
    font_size=15,
    font_name='Calibri',
    fill_params=cyan_title_pattern
)
```

If you need to declare multiple styles, and many of them are very similar, you
can use the `clone_and_modify` function to streamline the process. For example,
if you need to create blue_title_style, red_title_style, and green_title_style,
the straightforward method involves creating each one individually:

```python
blue_title_style = CustomStyle(
    font_bold=True,
    font_color='0000ff',
    font_size=15
)
red_title_style = CustomStyle(
    font_bold=True,
    font_color='ff0000',
    font_size=15
)
green_title_style = CustomStyle(
    font_bold=True,
    font_color='00ff00',
    font_size=15
)
```

However, with the clone_and_modify method, you can create them more efficiently:

```python
blue_title_style = CustomStyle(
    font_bold=True,
    font_color='0000ff',
    font_size=15
)
red_title_style = blue_title_style.clone_and_modify(font_color='ff0000')
green_title_style = blue_title_style.clone_and_modify(font_color='00ff00')
```

This approach is especially helpful when there are many parameters to input in
a CustomStyle instance.

## Advanced Usage

In scenarios where you want to establish a collection of reusable styles for your Excel documents or if you have a variety of different Excel outputs, you can define a base class containing multiple custom styles and then inherit from this base class in your Excel class.

The following example demonstrates how to achieve this using the openpyxl_style_writer library:


```python
from openpyxl_style_writer import CustomStyle, RowWriter


class BaseExcelWriter(RowWriter):
    blue_font = {
        'color': '0000ff',
        'bold': True,
        'size': 15,
    }
    cyan_title_pattern = {
        'patternType': 'solid',
        'fgColor': '00ffff	'
    }
    blue_font_style = CustomStyle(font_params=blue_font)
    cyan_fill_style = CustomStyle(fill_params=cyan_title_pattern)
    # Clone a existed style and modify font params
    cyan_fill_with_blue_font = cyan_fill_style.cloned_and_modify(
        font_params=blue_font
    )


class ExampleExcel(BaseExcelWriter):
    first_row = ['Apple', 'Banana', 'Cat']

    def create(self, file_name='output.xlsx'):
        self.create_sheet('ExampleSheet', protection=True)
        self.row_append_list(self.first_row, style=self.cyan_fill_style)
        self.create_row()
        for idx, _ in enumerate(self.first_row):
            self.row_append(idx, style=self.blue_font_style)
        self.row_append(
            'Clone_and_modify_cell',
            style=self.cyan_fill_with_blue_font
        )
        self.create_row()
        self.save(file_name)


class ExampleExcel2(BaseExcelWriter):

    def create(self, file_name='output.xlsx'):
        # create a content of excel2...

if __name__ == '__main__':
    example = ExampleExcel()
    example.create('example.xlsx')
```

## List of Key words in openpyxl_style_writer

This is a list of the key words in openpyxl_style_writer and how it map to the attributes of openpyxl

<div align='center'>

| **class**     |       Key           |  datatype   |          **map to**                                                   |
| :------------ | ------------------- | ----------- | :-------------------------------------------------------------------  |
| **font**      | font_size           | int         | openpyxl.styles.Font.size                                             |
|               | font_name           | str         | openpyxl.styles.Font.name                                             |
|               | font_bold           | bool        | openpyxl.styles.Font.bold                                             |
|               | font_italic         | bool        | openpyxl.styles.Font.italic                                           |
|               | font_underline      | str         | openpyxl.styles.Font.underline                                        |
|               | font_strike         | bool        | openpyxl.styles.Font.strike                                           |
|               | font_vertAlign      | str         | openpyxl.styles.Font.vertAlign                                        |
|               | font_color          | str         | openpyxl.styles.Font.color                                            |
| **fill**      | fill_color          | str         | openpyxl.styles.PatternFill.color                                     |
| **alignment** | ali_horizontal      | str         | openpyxl.styles.Alignment.color                                       |
|               | ali_vertical        | str         | openpyxl.styles.Alignment.color                                       |
|               | ali_wrap_text       | str         | openpyxl.styles.Alignment.color                                       |
| **border**    | border_style_top    | str         | openpyxl.styles.Border.top with openpyxl.styles.Side.border_style     |
|               | border_style_right  | str         | openpyxl.styles.Border.right with openpyxl.styles.Side.border_style   |
|               | border_style_left   | str         | openpyxl.styles.Border.left with openpyxl.styles.Side.border_style    |
|               | border_style_bottom | str         | openpyxl.styles.Border.bottom with openpyxl.styles.Side.border_style  |
|               | border_color_top    | str         | openpyxl.styles.Border.top with openpyxl.styles.Side.color            |
|               | border_color_right  | str         | openpyxl.styles.Border.right with openpyxl.styles.Side.color          |
|               | border_color_left   | str         | openpyxl.styles.Border.left with openpyxl.styles.Side.color           |
|               | border_color_bottom | str         | openpyxl.styles.Border.bottom with openpyxl.styles.Side.color         |

</div>
