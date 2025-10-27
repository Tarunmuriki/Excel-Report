from openpyxl.styles import PatternFill, Font, Alignment

def format_currency(value):
    """Format numeric value as currency string."""
    return f"${value:,.2f}" if isinstance(value, (int, float)) else value

def apply_excel_styles(worksheet):
    """Apply consistent styling to Excel worksheet."""
    # Style header row
    header_fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
    header_font = Font(bold=True)
    
    for cell in next(worksheet.rows):
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    
    # Adjust column widths
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        
        adjusted_width = (max_length + 2) * 1.2
        worksheet.column_dimensions[column_letter].width = min(adjusted_width, 50)