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
            ('font_size', 11),
            ('font_name', 'Calibri'),
            ('font_bold', False),
            ('font_italic', False),
            ('font_underline', 'none'),
            ('font_strike', False),
            ('font_vertAlign', None),
            ('font_color', '000000'),
            # fill attr
            ('fill_pattern', 'solid'),
            ('fill_color', 'ffffff'),
            # alignment attr
            ('ali_horizontal', None),
            ('ali_vertical', 'bottom'),
            ('ali_wrap_text', False),
            ('ali_text_rotation', 0),
            ('ali_shrink_to_fit', False),
            ('ali_indent', 0),
            # border attr
            ('border_style_top', 'thin'),
            ('border_style_right', 'thin'),
            ('border_style_left', 'thin'),
            ('border_style_bottom', 'thin'),
            ('border_color_top', 'C0C0C0'),
            ('border_color_right', 'C0C0C0'),
            ('border_color_left', 'C0C0C0'),
            ('border_color_bottom', 'C0C0C0'),
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


@pytest.mark.style
class TestStylesArgs:
    @pytest.mark.parametrize(
        'font_settings, expected_font_size, expected_font_bold,'
        + 'expected_font_color, expected_font_name',
        [
            (
                {
                    'font_size': 16,
                    'font_bold': True,
                    'font_color': 'cccccc',
                    'font_name': 'Calibri',
                },
                16,
                True,
                'cccccc',
                'Calibri',
            ),
            (
                {
                    'font_size': 19,
                    'font_bold': False,
                    'font_color': '00ff00',
                    'font_name': 'Arial',
                },
                19,
                False,
                '00ff00',
                'Arial',
            ),
        ],
    )
    def test_font_style(
        self,
        font_settings,
        expected_font_size,
        expected_font_bold,
        expected_font_color,
        expected_font_name,
    ):
        DefaultStyle.set_default(**font_settings)

        assert DefaultStyle.font_size == expected_font_size
        assert DefaultStyle.font_bold == expected_font_bold
        assert DefaultStyle.font_name == expected_font_name
        assert DefaultStyle.font_color == expected_font_color

        custom = CustomStyle(**font_settings)

        assert custom.font_size == expected_font_size
        assert custom.font_bold == expected_font_bold
        assert custom.font_name == expected_font_name
        assert custom.font_color == expected_font_color

    @pytest.mark.parametrize(
        'fill_settings, expected_fill_color',
        [
            (
                {'fill_color': 'cccccc'},
                'cccccc',
            ),
            (
                {'fill_color': 'ff0000'},
                'ff0000',
            ),
        ],
    )
    def test_fill_style(
        self,
        fill_settings,
        expected_fill_color,
    ):
        DefaultStyle.set_default(**fill_settings)
        assert DefaultStyle.fill_color == expected_fill_color

        custom = CustomStyle(**fill_settings)
        assert custom.fill_color == expected_fill_color

    @pytest.mark.parametrize(
        'border_settings, expected_border_top, expected_border_right,'
        + 'expected_border_left, expected_border_bottom',
        [
            (
                {
                    'border_style_top': None,
                    'border_style_right': 'dashDot',
                    'border_style_left': 'dashed',
                    'border_style_bottom': 'double',
                    'border_color_top': 'ff000000',
                    'border_color_right': '00000000',
                    'border_color_left': '00993366',
                    'border_color_bottom': '00FF99CC',
                },
                (None, 'ff000000'),
                ('dashDot', '00000000'),
                ('dashed', '00993366'),
                ('double', '00FF99CC'),
            ),
            (
                {
                    'border_style_top': 'hair',
                    'border_style_right': 'medium',
                    'border_style_left': 'mediumDashDot',
                    'border_style_bottom': 'mediumDashDotDot',
                    'border_color_top': '00080000',
                    'border_color_right': '00808000',
                    'border_color_left': '00CCFFCC',
                    'border_color_bottom': '00FFFF00',
                },
                ('hair', '00080000'),
                ('medium', '00808000'),
                ('mediumDashDot', '00CCFFCC'),
                ('mediumDashDotDot', '00FFFF00'),
            ),
        ],
    )
    def test_border_style(
        self,
        border_settings,
        expected_border_top,
        expected_border_right,
        expected_border_left,
        expected_border_bottom,
    ):
        DefaultStyle.set_default(**border_settings)
        assert (DefaultStyle.border_style_top, DefaultStyle.border_color_top) == expected_border_top
        assert (
            DefaultStyle.border_style_right,
            DefaultStyle.border_color_right,
        ) == expected_border_right
        assert (
            DefaultStyle.border_style_left,
            DefaultStyle.border_color_left,
        ) == expected_border_left
        assert (
            DefaultStyle.border_style_bottom,
            DefaultStyle.border_color_bottom,
        ) == expected_border_bottom

        custom = CustomStyle(**border_settings)
        assert (custom.border_style_top, custom.border_color_top) == expected_border_top
        assert (custom.border_style_right, custom.border_color_right) == expected_border_right
        assert (custom.border_style_left, custom.border_color_left) == expected_border_left
        assert (custom.border_style_bottom, custom.border_color_bottom) == expected_border_bottom

    @pytest.mark.parametrize(
        'ali_settings, expected_ali_horizontal, expected_ali_vertical, expected_ali_wrap_text',
        [
            (
                {
                    'ali_horizontal': 'general',
                    'ali_vertical': 'bottom',
                    'ali_wrap_text': False,
                },
                'general',
                'bottom',
                False,
            ),
            (
                {
                    'ali_horizontal': 'distributed',
                    'ali_vertical': 'top',
                    'ali_wrap_text': True,
                },
                'distributed',
                'top',
                True,
            ),
        ],
    )
    def test_alignment_style(
        self,
        ali_settings,
        expected_ali_horizontal,
        expected_ali_vertical,
        expected_ali_wrap_text,
    ):
        DefaultStyle.set_default(**ali_settings)
        assert DefaultStyle.ali_horizontal == expected_ali_horizontal
        assert DefaultStyle.ali_vertical == expected_ali_vertical
        assert DefaultStyle.ali_wrap_text == expected_ali_wrap_text

        custom = CustomStyle(**ali_settings)
        assert custom.ali_horizontal == expected_ali_horizontal
        assert custom.ali_vertical == expected_ali_vertical
        assert custom.ali_wrap_text == expected_ali_wrap_text


@pytest.mark.parametrize(
    'original_color, modified_color, expected_original_color, expected_modified_color',
    [
        ('ff0000', '0000ff', '00ff0000', '000000ff'),
        ('00ff00', 'ff00ff', '0000ff00', '00ff00ff'),
        ('0000ff', 'ffff00', '000000ff', '00ffff00'),
    ],
)
def test_clone_and_modify(
    original_color,
    modified_color,
    expected_original_color,
    expected_modified_color,
):
    original_style = CustomStyle(font_size=19, font_color=original_color, font_bold=True)

    assert original_style.font.size == 19
    assert original_style.font.color.rgb == expected_original_color
    assert original_style.font.bold is True

    modified_style = original_style.clone_and_modify(font_color=modified_color)

    assert modified_style.font.size == 19
    assert modified_style.font.color.rgb == expected_modified_color
    assert modified_style.font.bold is True
