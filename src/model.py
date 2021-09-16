
class Model:
    files = []
    output_workbook = None
    path_to_save = None
    in_filetypes = (    
                        ("Excel Workbook", "*.xlsx"),
                        ("Excel Macro-Enabled Workbook", "*.xlsm"),
                        ("Excel Macro-Enabled Workbook Template", "*.xltm"),
                        ("Excel Spreadsheet Template", "*.xltx"),
                        # ("OpenDocument Spreadsheet", "*.ods"),
                        # ("CSV (Comma delimited)", "*.csv"),
                        # ("Excel 97-2003 Workbook", "*.xls"),
                        # ("Excel Binary Workbook", "*.xlsb")
                    )

    out_filetypes = (   
                        ("Excel Workbook", "*.xlsx"),
                        # ("OpenDocument Spreadsheet", "*.ods"),
                        # ("CSV (Comma delimited)", "*.csv"),
                    )
