import pytest
from openpyxl.styles import Side
from openpyxl.styles import Protection
from openpyxl_style_writer import DefaultStyle, CustomStyle


@pytest.mark.style
class TestStyles:
    @pytest.mark.parametrize(
        'attr, expected',
        [
            # params attr
            ('font_params', None),
            ('fill_params', None),
            ('ali_params', None),
            ('border_params', None),
            # font attr
            ('font_size', 14),
            ('font_name', 'Calibri'),
            ('font_bold', False),
            ('font_italic', False),
            ('font_underline', 'none'),
            ('font_strike', False),
            ('font_vertAlign', None),
            ('font_color', '000000'),
            # fill attr
            ('fill_pattern', 'solid'),
            ('fill_color', 'fcfcfc'),
            # alignment attr
            ('ali_horizontal', 'center'),
            ('ali_vertical', 'center'),
            ('ali_wrap_text', False),
            ('ali_text_rotation', 0),
            ('ali_shrink_to_fit', False),
            ('ali_indent', 0),
            # border attr
            ('border_style_top', None),
            ('border_style_right', None),
            ('border_style_left', None),
            ('border_style_bottom', None),
            ('border_color_top', 'ff000000'),
            ('border_color_right', 'ff000000'),
            ('border_color_left', 'ff000000'),
            ('border_color_bottom', 'ff000000'),
            # protect attr
            ('protect', False),
            ('protection', Protection(locked=False)),
            # format attr
            ('number_format', 'General'),
        ],
    )
    def test_default_and_custom_style_creation(self, attr, expected):
        default_style = DefaultStyle()
        assert getattr(default_style, attr) == expected
        custom_style_without_setting = CustomStyle()
        assert getattr(custom_style_without_setting, attr) == expected

    @pytest.mark.parametrize(
        'font_params, expected_font_size, expected_font_bold',
        [
            ({'size': 16, 'bold': True}, 16, True),
            ({'size': 12, 'bold': False}, 12, False),
        ],
    )
    def test_init_and_apply_settings_with_font_params(
        self,
        font_params,
        expected_font_size,
        expected_font_bold,
    ):
        custom_style = CustomStyle(font_params=font_params)
        assert custom_style.font.size == expected_font_size
        assert custom_style.font.bold == expected_font_bold

    @pytest.mark.parametrize(
        'fill_params, expected_fill_pattern, expected_fill_color',
        [
            ({'patternType': 'solid', 'fgColor': 'ffffff'}, 'solid', '00ffffff'),
            ({'patternType': 'gray125', 'fgColor': '999999'}, 'gray125', '00999999'),
        ],
    )
    def test_apply_settings_with_fill_params(
        self,
        fill_params,
        expected_fill_pattern,
        expected_fill_color,
    ):
        custom_style = CustomStyle(fill_params=fill_params)

        assert custom_style.fill.patternType == expected_fill_pattern
        assert custom_style.fill.fgColor.rgb == expected_fill_color

    @pytest.mark.parametrize(
        'ali_params, expected_horizontal, expected_vertical, expected_wrap_text',
        [
            ({'horizontal': 'center', 'vertical': 'top', 'wrap_text': True}, 'center', 'top', True),
            (
                {'horizontal': 'right', 'vertical': 'bottom', 'wrap_text': False},
                'right',
                'bottom',
                False,
            ),
        ],
    )
    def test_apply_settings_with_alignment(
        self,
        ali_params,
        expected_horizontal,
        expected_vertical,
        expected_wrap_text,
    ):
        custom_style = CustomStyle(ali_params=ali_params)
        assert custom_style.ali.horizontal == expected_horizontal
        assert custom_style.ali.vertical == expected_vertical
        assert custom_style.ali.wrap_text == expected_wrap_text

    @pytest.mark.parametrize(
        'border_params, expected_border_style_top, expected_border_style_right,'
        + 'expected_border_style_bottom, expected_border_style_left'
        + ', expected_border_color_top, expected_border_color_right,'
        + 'expected_border_color_bottom, expected_border_color_left',
        [
            (
                {
                    'left': Side(style='dotted', color='cccccc'),
                    'right': Side(style='medium', color='000000'),
                    'top': Side(style='thin', color='ff0000'),
                    'bottom': Side(style='thick', color='00ff00'),
                },
                'thin',
                'medium',
                'thick',
                'dotted',
                '00ff0000',
                '00000000',
                '0000ff00',
                '00cccccc',
            ),
            (
                {
                    'left': Side(style='medium', color='cccccc'),
                    'right': Side(style='thin', color='cccccc'),
                    'top': Side(style='double', color='cccccc'),
                    'bottom': Side(style='dashed', color='cccccc'),
                },
                'double',
                'thin',
                'dashed',
                'medium',
                '00cccccc',
                '00cccccc',
                '00cccccc',
                '00cccccc',
            ),
        ],
    )
    def test_apply_settings_with_border_params(
        self,
        border_params,
        expected_border_style_top,
        expected_border_style_right,
        expected_border_style_bottom,
        expected_border_style_left,
        expected_border_color_top,
        expected_border_color_right,
        expected_border_color_bottom,
        expected_border_color_left,
    ):
        custom_style = CustomStyle(border_params=border_params)
        assert custom_style.border.top.style == expected_border_style_top
        assert custom_style.border.right.style == expected_border_style_right
        assert custom_style.border.bottom.style == expected_border_style_bottom
        assert custom_style.border.left.style == expected_border_style_left

        assert custom_style.border.top.color.rgb == expected_border_color_top
        assert custom_style.border.right.color.rgb == expected_border_color_right
        assert custom_style.border.bottom.color.rgb == expected_border_color_bottom
        assert custom_style.border.left.color.rgb == expected_border_color_left

    @pytest.mark.parametrize(
        'number_format, expected_number_format',
        [
            ('General', 'General'),
            ('0.00', '0.00'),
            ('#,##0', '#,##0'),
        ],
    )
    def test_apply_settings_with_number_format(self, number_format, expected_number_format):
        custom_style = CustomStyle(number_format=number_format)
        assert custom_style.number_format == expected_number_format

    @pytest.mark.parametrize(
        'protect, expected_protection',
        [
            (True, True),
            (False, False),
        ],
    )
    def test_apply_settings_with_protect(self, protect, expected_protection):
        custom_style = CustomStyle(protect=protect)
        assert custom_style.protect == expected_protection
        assert custom_style.protection == Protection(locked=expected_protection)
