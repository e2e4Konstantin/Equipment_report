# https://htmlcolorcodes.com/
# https://htmlcolorcodes.com/color-chart/

from openpyxl.styles import Color, Font, PatternFill, Border, Side

side_none = Side(border_style=None)

headers = {
    'equipments_main': ['номер', 'место', 'шифр', 'описание', 'измерение', 'статистика', 'флаг', 'примечание', 'атрибуты/параметры', ''],

    'option': ['от', 'до', 'ед.изм.', 'шаг', 'тип',],
    'attribute': ['атрибуты',],
    'table': ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
}

width_columns = {'A': 5, 'B': 4, 'C': 12, 'D': 100, 'E': 10, 'F': 7, 'G': 4, 'H': 8, 'I': 10, 'J': 10, 'K': 10, 'L': 10, 'M': 10, 'N': 10}


items_styles = {
                'equipment': {
                    'font': Font(name='Calibri', color=Color(rgb='00000000'), size=8, bold=False),
                    'fill': PatternFill(patternType="solid", fgColor=Color(rgb='00FFFFFF')),
                    'border': Border(top=Side(border_style=None), bottom=Side(border_style='thin', color=Color(rgb='00E5E7E9')))
                },

                'line': {
                    'font': Font(name='Calibri', color=Color(rgb='00D35400'), size=8, bold=False),
                    'fill': PatternFill(patternType="solid", fgColor=Color(rgb='000DF04E')),
                    'border': Border(top=Side(border_style='thin', color=Color(rgb='00A0A0A0')), bottom=Side(border_style='thin', color=Color(rgb='00A0A0A0')))
                },
                'table': {
                    'font': Font(name='Calibri', color=Color(rgb='0034495E'), size=8, bold=False),
                    'fill': PatternFill(patternType="solid", fgColor=Color(rgb='00FCFBFA')),
                    'border_up': Border(top=Side(border_style='thin', color=Color(rgb='00A0A0A0'))),
                    'border_down': Border(bottom=Side(border_style='thin', color=Color(rgb='00A0A0A0')))
                },
                'main_equipments_header': {
                    'font': Font(name='Calibri', color=Color(rgb='0034495E'), size=8, bold=True),
                    'fill': PatternFill(patternType="solid", fgColor=Color(rgb='00E4DFEC')),
                    'border': Border(top=Side(border_style='thin', color=Color(rgb='00A0A0A0')), bottom=Side(border_style='thin', color=Color(rgb='00A0A0A0')))
                },

                'attributes_header': {
                    'font': Font(name='Calibri', color=Color(rgb='0034495E'), size=8, bold=False),
                    'fill': PatternFill(patternType="solid", fgColor=Color(rgb='00F7F7F8')),
                    'border': Border(
                                    top=Side(border_style='thin', color=Color(rgb='00A0A0A0')),
                                    bottom=Side(border_style='thin', color=Color(rgb='00A0A0A0')),
                                    left=Side(border_style='thin', color=Color(rgb='00A0A0A0')),
                                    right=Side(border_style='thin', color=Color(rgb='00A0A0A0'))
                                    )
                },
                'attributes_volume': {
                    'font': Font(name='Calibri', color=Color(rgb='0034495E'), size=8, bold=True),
                    'fill': PatternFill(patternType="solid", fgColor=Color(rgb='00F7F7F8')),
                    'border': Border(
                                    top=Side(border_style='thin', color=Color(rgb='00A0A0A0')),
                                    bottom=Side(border_style='thin', color=Color(rgb='00A0A0A0')),
                                    left=Side(border_style='thin', color=Color(rgb='00A0A0A0')),
                                    right=Side(border_style='thin', color=Color(rgb='00A0A0A0'))
                                    )
                },

                'parameter_title': {
                    'font': Font(name='Calibri', color=Color(rgb='0034495E'), size=8, bold=True),
                    'fill': PatternFill(patternType="solid", fgColor=Color(rgb='00EBF1DE')),    # 00DFFF00
                    'border': Border(
                                    top=Side(border_style='thin', color=Color(rgb='00A0A0A0')),
                                    bottom=Side(border_style='thin', color=Color(rgb='00A0A0A0')),
                                    left=Side(border_style='thin', color=Color(rgb='00A0A0A0')),
                                    right=Side(border_style='thin', color=Color(rgb='00A0A0A0'))
                                    )
                },
                'parameter_table': {
                    'font': Font(name='Calibri', color=Color(rgb='0034495E'), size=8, bold=False),
                    'fill': PatternFill(patternType="solid", fgColor=Color(rgb='00EBF1DE')), # 006495ED
                    'border': Border(
                                    top=Side(border_style='thin', color=Color(rgb='00A0A0A0')),
                                    bottom=Side(border_style='thin', color=Color(rgb='00A0A0A0')),
                                    left=Side(border_style='thin', color=Color(rgb='00A0A0A0')),
                                    right=Side(border_style='thin', color=Color(rgb='00A0A0A0'))
                                    )
                },


}

basic_colors = {
    'grey': Font(name='Calibri', color=Color(rgb='0099A3A4'), size=8, bold=False),
    'indigo':	Font(name='Calibri', color=Color(rgb='004B0082'), size=8, bold=False),
    'dark_red':	Font(name='Calibri', color=Color(rgb='008B0000'), size=8, bold=False),
    'black':	Font(name='Calibri', color=Color(rgb='00000000'), size=8, bold=False),
    'dark_blue_bold':	Font(name='Calibri', color=Color(rgb='0000008B'), size=8, bold=True),
}


# own_border = Border(left=Side(border_style=None, color='FF000000'),
#                     right=Side(border_style=None, color='FF000000'),
#                     top=Side(border_style=None, color='FF000000'),
#                     bottom=Side(border_style=None, color='FF000000'),
#                     diagonal=Side(border_style=None, color='FF000000'),
#                     diagonal_direction=0,
#                     outline=Side(border_style=None, color='FF000000'),
#                     vertical=Side(border_style=None, color='FF000000'),
#                     horizontal=Side(border_style=None, color='FF000000')
#                     )






