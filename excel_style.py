from openpyxl.styles import Border, Side, Alignment
from openpyxl.styles import PatternFill, Font
from datetime import datetime

def border_alignCenter():
    border = Border(left=Side(border_style="thin", color="000000"),
                    right=Side(border_style="thin", color="000000"),
                    top=Side(border_style="thin", color="000000"),
                    bottom=Side(border_style="thin", color="000000"))
    alignment_center = Alignment(horizontal="center", vertical="center") 
    return border, alignment_center

def BoldFont():
    return Font(bold=True)

def FillColor(color_code):
    return PatternFill(start_color=color_code, end_color=color_code, fill_type="solid")