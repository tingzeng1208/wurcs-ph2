import os
import sys
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, '..'))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

from services.DBManager import DBManager
from data.db_data import db_data
from utils.utility import to_str

class CreateReports:
    
    def __init__(self, year, db_manager: DBManager= None):
        if db_manager is None:
            db_manager = DBManager()
        self.o_db = db_manager
        self.db_data = db_manager.db_data
        self.current_year = year
        db_manager.db_data.s_current_year = year
        self.current_railroad = None
        self.current_df = None
        self.sTitle_RR_YEAR = ""

        # Initialize all dt*** variables to empty DataFrames with None
        self.dtTitles = pd.DataFrame(None)
        
        self.dtFootnotes = pd.DataFrame(None)
        self.dtAValue = pd.DataFrame(None)
        self.dtAValue0_RR = pd.DataFrame(None)
        self.dtAValueRegion_RR = pd.DataFrame(None)
        self.dtPriceIndexes = pd.DataFrame(None)
        self.dtPriceIndexReferences = pd.DataFrame(None)
        self.dtRegression = pd.DataFrame(None)
        self.dtRegressionDependentVariable = pd.DataFrame(None)
        self.dtRegressionDistribution = pd.DataFrame(None)
        self.dtDefinitions = pd.DataFrame(None)
        self.dtCarTypeStatistics = pd.DataFrame(None)
        self.dtCarTypeStatisticsPart2 = pd.DataFrame(None)
        self.dtCarTypeStatisticsPart3 = pd.DataFrame(None)
        self.dtLineSourceText = pd.DataFrame(None)
        self.dtDataDictionary = pd.DataFrame(None)
        self.dtECodes = pd.DataFrame(None)
        self.rr_icc = None
        self.s_output_folder = os.path.join(parent_dir, "..", "reports")
    
    def initialize_report_for_a_railroad(self, rr_no):
        self.dtAValueRegion_RR = self.o_db.get_a_value_region_rr(rr_no, self.current_year)
        print("Loaded dtAValueRegion_RR. number of rows:", len(self.dtAValueRegion_RR))
        self.dtPriceIndexes = self.o_db.get_price_indexes(rr_no, str(self.current_year))
        print("Loaded dtPriceIndexes. number of rows:", len(self.dtPriceIndexes))        
        self.dtAValue = self.o_db.get_a_value(rr_no, self.current_year)
        print("Loaded dtAValue, number of rows:", len(self.dtAValue))
        self.dtCarTypeStatistics = self.o_db.get_car_type_statistics(rr_no)        
        print("Loaded dtCarTypeStatistics, number of rows:", len(self.dtCarTypeStatistics))
        self.dtCarTypeStatisticsPart2 = self.o_db.get_car_type_statistics_part2(rr_no)
        print("Loaded dtCarTypeStatisticsPart2, number of rows:", len(self.dtCarTypeStatisticsPart2))

    def initialize_report_format(self):
        print("Initializing report format...")
        self.dtTitles = self.o_db.get_custom_data("REPORT_TITLES")
        print("Loaded dtTitles. number of rows:", len(self.dtTitles))
        self.dtFootnotes = self.o_db.get_custom_data("REPORT_FOOTNOTES")
        print("Loaded dtFootnotes. number of rows:", len(self.dtFootnotes))
        self.dtAValue0_RR = self.o_db.get_a_value0_rr(self.current_year)
        print("Loaded dtAValue0_RR. number of rows:", len(self.dtAValue0_RR))
        self.dtPriceIndexReferences = self.o_db.get_custom_data("PRICE_INDEX_REFERENCE")
        print("Loaded dtPriceIndexReferences. number of rows:", len(self.dtPriceIndexReferences))
        self.dtRegression = self.o_db.get_custom_data("REGR")
        print("Loaded dtRegression. number of rows:", len(self.dtRegression))
        self.dtRegressionDependentVariable = self.o_db.get_custom_data("REGR_DEP_VAR")
        print("Loaded dtRegressionDependentVariable. number of rows:", len(self.dtRegressionDependentVariable))
        self.dtRegressionDistribution = self.o_db.get_custom_data("REGR_DISTRIB")
        print("Loaded dtRegressionDistribution. number of rows:", len(self.dtRegressionDistribution))
        self.dtDefinitions = self.o_db.get_custom_data("URCSDEF")
        print("Loaded dtDefinitions. number of rows:", len(self.dtDefinitions))
        self.dtCarTypeStatisticsPart3 = self.o_db.get_car_type_statistics_part3()
        print("Loaded dtCarTypeStatisticsPart3. number of rows:", len(self.dtCarTypeStatisticsPart3))
        self.dtLineSourceText = self.o_db.get_line_source_text()
        print("Loaded dtLineSourceText. number of rows:", len(self.dtLineSourceText))
        self.dtDataDictionary = self.o_db.get_data_dictionary(self.current_year)
        print("Loaded dtDataDictionary. number of rows:", len(self.dtDataDictionary))
        self.dtECodes = self.o_db.get_custom_data("ECODES")
        print("Loaded dtECodes. number of rows:", len(self.dtECodes))

    def print_initialized_dataframes(self):
        print("dtTitles:")
        print(self.dtTitles)
        print("dtFootnotes:")
        print(self.dtFootnotes)
        print("dtAValue:")
        print(self.dtAValue)
        print("dtAValue0_RR:")
        print(self.dtAValue0_RR)
        print("dtAValueRegion_RR:")
        print(self.dtAValueRegion_RR)
        print("dtPriceIndexes:")
        print(self.dtPriceIndexes)
        print("dtPriceIndexReferences:")
        print(self.dtPriceIndexReferences)
        print("dtRegression:")
        print(self.dtRegression)
        print("dtRegressionDependentVariable:")
        print(self.dtRegressionDependentVariable)
        print("dtRegressionDistribution:")
        print(self.dtRegressionDistribution)
        print("dtDefinitions:")
        print(self.dtDefinitions)
        print("dtCarTypeStatistics:")
        print(self.dtCarTypeStatistics)
        print("dtCarTypeStatisticsPart2:")
        print(self.dtCarTypeStatisticsPart2)
        print("dtCarTypeStatisticsPart3:")
        print(self.dtCarTypeStatisticsPart3)
        print("dtLineSourceText:")
        print(self.dtLineSourceText)
        print("dtDataDictionary:")
        print(self.dtDataDictionary)
        print("dtECodes:")
        print(self.dtECodes)

    def setup_report(self, output_folder, current_year, rr_no, rricc, rr_name, rr_longname,
                    null_string_replace, run_years, capital_cost, variability_flow, account76,
                    account80, account90, ssac, create_log, db):
        """
        Sets up report parameters and returns them as a dictionary (no class used).
        Equivalent to the C# constructor logic.
        """
        from datetime import datetime
        sFileName = f"{output_folder}\\{rr_name}{current_year}"
        sLogFileName = f"{output_folder}\\Log\\{rr_name}{current_year}.log"
        sTitle_RR_YEAR = f"{rr_name} {current_year} Run: {datetime.now().strftime('%m/%d/%Y %I:%M:%S %p')}"
        return {
            "oDB": db,
            "sFileName": sFileName,
            "sLogFileName": sLogFileName,
            "bCreateLog": create_log,
            "sNullValueString": null_string_replace,
            "iRunYears": run_years,
            "iRailroadNumber": rr_no,
            "iRRICC": rricc,
            "sRailRoadShortName": rr_name,
            "sRailRoadName": rr_longname,
            "bCapitalCost": capital_cost,
            "bVarialabilityFlow": variability_flow,
            "bAccount76": account76,
            "bAccount80": account80,
            "bAccount90": account90,
            "bSSAC": ssac,
            "iCurrentYear": int(current_year),
            "sTitle_RR_YEAR": sTitle_RR_YEAR
        }
        
    def remove_default_sheet(self, wb):
        # Remove the default sheet if it exists
        default_sheetnames = ["Sheet", "Sheet1"]
        for sheetname in default_sheetnames:
            if sheetname in wb.sheetnames:
                std = wb[sheetname]
                wb.remove(std)

    def create_workbook(self, workbookname, path):

        if not os.path.exists(path):
            os.makedirs(path)
        full_path = os.path.join(path, workbookname)
        wb = Workbook()
        # The default sheet "Sheet" is created. It will be removed later.
        wb.save(full_path)
        return wb, full_path

    def add_worksheet(self, workbook, sheet_name):
        """
        Adds a worksheet with the specified name to the provided openpyxl Workbook object.
        Args:
            workbook (openpyxl.Workbook): The workbook to add the worksheet to.
            sheet_name (str): The name of the worksheet to add.
        Returns:
            openpyxl.worksheet.worksheet.Worksheet: The created worksheet object.
        """
        return workbook.create_sheet(title=sheet_name)

    def create_reports(self):

        df = self.o_db.get_class1_rail_list()
        count = 0
        self.initialize_report_format()
        for row in df.itertuples(index=False):
            count += 1
            if (count>1):
                break
            self.current_df = row
            short_name = row.short_name
            self.initialize_report_for_a_railroad(row.rr_id)
            self.create_a_report(short_name)

    def set_cell_with_format_and_name(self, ws, wb, row, col, value, number_format, alignment, named_range):
        # Convert value to int if possible for number formatting
        try:
            value = int(value)
        except (ValueError, TypeError):
            pass
        cell = ws.cell(row=row, column=col, value=value)
        cell.number_format = number_format
        cell.alignment = alignment
        wb.defined_names[named_range] = DefinedName(name=named_range, attr_text=f"'{ws.title}'!${cell.column_letter}${cell.row}")
        return cell


    def create_a_report(self, railroad_shortname):

        workbookname = f"{railroad_shortname}-{self.o_db.db_data.s_current_year}_report.xlsx"
        wb, full_path = self.create_workbook(workbookname,self.s_output_folder)
        print(f"Workbook {workbookname} created at: {full_path}")
        
        # Configure workbook for formula calculation
        
        self.build_index_worksheet(wb)
        self.A1P1_worksheet(wb)
        self.A1P2A_worksheet(wb)
        # self.A1P2B_worksheet(wb) # Placeholder
        # self.A1P2C_worksheet(wb) # Placeholder

        self.remove_default_sheet(wb)
        wb.save(full_path)
        print(f"Workbook {workbookname} saved at: {full_path}")


    def build_index_worksheet(self, wb: Workbook):
        """
        Builds the 'INDEX' worksheet in the provided workbook, sets titles, and populates user input values with formatting.
        Args:
            workbook (openpyxl.Workbook): The workbook to add the worksheet to.
        Returns:
            openpyxl.worksheet.worksheet.Worksheet: The created worksheet object.
        """
        try:
            sTitle_WORKTABLE = "User Inputs and Index"
            sSheetTitle = "INDEX"
            ws = wb.create_sheet(title=sSheetTitle)

            # Set font for all cells
            default_font = Font(name="Consolas", size=10)
            for row in ws.iter_rows(min_row=1, max_row=18, min_col=1, max_col=2):
                for cell in row:
                    cell.font = default_font
            sTitle_RR_YEAR = f"{self.current_df.short_name} {self.o_db.db_data.s_current_year} Run: {datetime.now().strftime('%m/%d/%Y %I:%M:%S %p')}"
            self.sTitle_RR_YEAR = sTitle_RR_YEAR
            ws.cell(row=1, column=1, value=sTitle_RR_YEAR)
            ws.cell(row=1, column=1).font = Font(name="Consolas", size=10, italic=True)

            # Worktable title
            ws.cell(row=2, column=1, value=sTitle_WORKTABLE)
            ws.cell(row=2, column=1).font = Font(name="Consolas", size=12, bold=True)

            # User Input Table header
            ws.cell(row=4, column=1, value="User Input values:")
            ws.cell(row=4, column=1).font = Font(name="Consolas", size=12, bold=True)
            ws.cell(row=4, column=1).alignment = Alignment(horizontal="left")
            ws.merge_cells("A4:B4")
            # Add border and double bottom border
            border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='double')
            )
            for cell in ws["A4:B4"]:
                for c in cell:
                    c.border = border

            # User input values
            user_inputs = [
                ("RRICC:", str(self.current_df.rricc), "RRICC"),
                ("Railroad Short Name:", str(self.current_df.short_name), "RR_SHORT_NAME"),
                ("Railroad Name:", str(self.current_df.name), "RR_NAME"),
                ("Railroad ID:", str(self.current_df.rr_id), "RR_ID"),
                ("Current Year:", str(self.o_db.db_data.s_current_year), "CURRENT_YEAR"),
                ("Run Years:", str(self.o_db.db_data.RUN_YEARS), "RUN_YEARS"),
                ("Embedded Cost of Capital for ROI:", "Y" if self.o_db.db_data.CapitalCost else "N", "EMBEDDED_COC"),
                ("100% Variability Flow Thru:", "Y" if self.o_db.db_data.VarialabilityFlow else "N", "FLOW_THRU_100PCT"),
                ("Include Account 76:", "Y" if self.o_db.db_data.Account76 else "N", "INCL_ACCT_76"),
                ("Include Account 80:", "Y" if self.o_db.db_data.Account80 else "N", "INCL_ACCT_80"),
                ("Include Account 90:", "Y" if self.o_db.db_data.Account90 else "N", "INCL_ACCT_90"),
                ("SSAC:", "Y" if self.o_db.db_data.SSAC else "N", "INCL_SSAC"),
                ("Replace Null Values With:", str(self.o_db.db_data.NullStringReplace), "NULL_VALUE"),
            ]
            for idx, (label, value, name) in enumerate(user_inputs, start=5):
                ws.cell(row=idx, column=1, value=label)
                cell = ws.cell(row=idx, column=2, value=value)
                cell.font = Font(name="Consolas", size=10)
                cell.alignment = Alignment(horizontal="right")
                ws.column_dimensions[cell.column_letter].width = 20
                # Add named range for this cell
                wb.defined_names[name] = DefinedName(name=name, attr_text=f"'{sSheetTitle}'!${cell.column_letter}${cell.row}")

            # Set column width for A2 (worktable title)
            ws.column_dimensions['A'].width = 33

            return ws

        except Exception as ex:
            print(f"Error in build_index_worksheet: {ex}")
            return None

    def WriteFirst3ColumnsAndPageLayout(self, ws, dtLineSourceText, dtFootnotes, sSheetTitle, sSelectStringRows, sSelectWorktable):
        # Set page layout
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.zoom = 70
        ws.oddFooter.left.text = "Page &P of &N"
        ws.print_title_rows = '$1:$7'
        ws.print_title_cols = 'A:C'

        # Write first 3 columns from dtLineSourceText
        iROW_COUNT = 8
        for _, drRow in dtLineSourceText[dtLineSourceText["rpt_sheet"] == sSheetTitle].iterrows():
            ws.cell(row=iROW_COUNT, column=1, value=str(drRow["line"]))
            ws.cell(row=iROW_COUNT, column=2, value=str(drRow["code"]))
            ws.cell(row=iROW_COUNT, column=3, value=str(drRow["identification"]))
            iROW_COUNT += 1

        # Leave two rows, then write footnotes
        iROW_COUNT += 2
        for _, drRow in dtFootnotes[dtFootnotes["worktable"] == sSheetTitle[:2]][dtFootnotes["part"] == sSheetTitle[sSheetTitle.index('P')+1:sSheetTitle.index('P')+3]].iterrows():
            ws.cell(row=iROW_COUNT, column=2, value=str(drRow["no"]))
            ws.cell(row=iROW_COUNT, column=3, value=str(drRow["text"]))
            iROW_COUNT += 1
            
    def write_titles_and_column_headers(self, ws, dtTitles, sSelectWorktable, sTitle_RR_YEAR, sTitle_WORKTABLE, iColumnCount, sSheetTitle):
        # Set font for all cells
        for row in ws.iter_rows(min_row=1, max_row=7, min_col=1, max_col=iColumnCount):
            for cell in row:
                cell.font = Font(name="Consolas", size=10)

        # Title rows
        ws.cell(row=1, column=1, value=sTitle_RR_YEAR)
        ws.cell(row=1, column=1).font = Font(name="Consolas", size=10, italic=True)
        ws.cell(row=2, column=1, value=sTitle_WORKTABLE)
        ws.cell(row=2, column=1).font = Font(name="Consolas", size=12, bold=True)

        # Get matching title rows
        part_str = sSheetTitle[sSheetTitle.find('P')+1:]
        filtered_titles = dtTitles[(dtTitles["worktable"] == sSheetTitle[:2]) & (dtTitles["part"] == part_str)]

        for _, drTitles in filtered_titles.iterrows():
            # Worktable titles
            ws.cell(row=3, column=2, value=to_str(drTitles["title1"]))
            ws.cell(row=3, column=2).font = Font(bold=True)
            ws.cell(row=4, column=2, value=to_str(drTitles["title2"]))
            ws.cell(row=4, column=2).font = Font(bold=True)
            ws.cell(row=5, column=2, value=to_str(drTitles["title3"]))
            ws.cell(row=5, column=2).font = Font(bold=True)

            sWorktableColumn = ""
            # Column headers
            for iCurrentCol in range(1, iColumnCount + 1):
                sColumnName = f"rpt_col{iCurrentCol}"
                col_val = str(drTitles.get(sColumnName, ""))
                if "Source" in col_val:
                    sWorktableColumn = col_val.split(" ")[0]
                else:
                    if sWorktableColumn:
                        ws.cell(row=6, column=iCurrentCol, value=sWorktableColumn)
                        ws.cell(row=6, column=iCurrentCol).alignment = Alignment(horizontal="center")
                ws.cell(row=7, column=iCurrentCol, value=col_val.replace("|", "\n"))

            # Column header formatting
            # Calculate Excel column range string for formatting
            def excel_col(n):
                result = ""
                while n > 0:
                    n, rem = divmod(n - 1, 26)
                    result = chr(65 + rem) + result
                return result

            start_col = "A"
            end_col = excel_col(iColumnCount)
            sColumnRange = f"{start_col}7:{end_col}7"
            for row in ws[sColumnRange]:
                for cell in row:
                    cell.alignment = Alignment(horizontal="center", wrap_text=True)
                    cell.border = Border(bottom=Side(style='double'), left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'))
                    cell.font = Font(bold=True)
            # Set row height for header row
            ws.row_dimensions[7].height = 40
            # Set column width for header columns
            for iCurrentCol in range(1, iColumnCount + 1):
                col_letter = excel_col(iCurrentCol)
                ws.column_dimensions[col_letter].width = 25

            # Setup line, code, and identification columns
            ws.column_dimensions['A'].width = 8
            ws.column_dimensions['B'].width = 8
            ws.column_dimensions['C'].width = 50
            for col in ['A', 'B', 'C']:
                for cell in ws[col]:
                    cell.alignment = Alignment(horizontal="left")
            # SET PAGE LAYOUT (openpyxl syntax)
            ws.page_setup.orientation = "landscape"
            ws.page_setup.zoom = 70
            ws.oddFooter.left.text = "Page &P of &N"
            ws.print_title_rows = '1:7'
            ws.print_title_cols = 'A:C'

            # Write the line source text
            iROW_COUNT = 8

            for _, drRow in self.dtLineSourceText[self.dtLineSourceText["rpt_sheet"] == sSheetTitle].iterrows():
                # Write The Line Number
                ws.cell(row=iROW_COUNT, column=1, value=getattr(drRow, "line", getattr(drRow, "Line", None)))
                # Write Code
                ws.cell(row=iROW_COUNT, column=2, value="'" + str(getattr(drRow, "code", getattr(drRow, "Code", ""))))
                # Write Identification
                ws.cell(row=iROW_COUNT, column=3, value=str(getattr(drRow, "identification", getattr(drRow, "Identification", ""))))
                # increment iROW_COUNT
                iROW_COUNT += 1
            iROW_COUNT += 2
            # Write footnotes
            filtered_footnotes = self.dtFootnotes[(self.dtFootnotes["worktable"] == sSheetTitle[:2]) & (self.dtFootnotes["part"] == part_str)]
            for drRow in filtered_footnotes.itertuples(index=False):
                # Write The Footnote number
                ws.cell(row=iROW_COUNT, column=2, value=str(getattr(drRow, "no", getattr(drRow, "No", ""))))
                # Write Code
                ws.cell(row=iROW_COUNT, column=3, value="'" + str(getattr(drRow, "text", getattr(drRow, "Text", ""))))
                # increment iROW_COUNT
                iROW_COUNT += 1

    def scrub_year(self, s: str, current_year: int) -> str:
        s = s.replace("#Y#", str(current_year))
        s = s.replace("#Y-1#", str(current_year - 1))
        s = s.replace("#Y-2#", str(current_year - 2))
        s = s.replace("#Y-3#", str(current_year - 3))
        s = s.replace("#Y-4#", str(current_year - 4))
        s = s.replace("#Y-5#", str(current_year - 4))
        return s
    
    
    def add_formula_as_text(self, ws, row, col, formula):
        """
        Writes a formula as a string (not executable) in the specified cell.
        This is useful for displaying the formula as text in Excel, not as a calculated value.
        """
        cell = ws.cell(row=row, column=col, value=f"'{formula}")
        cell.alignment = Alignment(horizontal="left")
    
    def A1P1_worksheet(self, wb, index=2):
        try:
            sTitle_WORKTABLE = "WORKTABLE A1 PART 1"
            iColumnCount = 21
            sSheetTitle = "A1P1"
            iLineNumberOffset = 93

            dtaValue = self.dtAValue
            dtaValue0_RR = self.dtAValue0_RR
            dtLineSourceText = self.dtLineSourceText
            iCurrentYear = int(self.current_year)

            # Update status
            print(f"Processing {sSheetTitle}")

            # Select worktable and string rows
            sSelectWorktable = f"Worktable = '{sSheetTitle[:2]}' And Part = '{sSheetTitle[sSheetTitle.index('P')+1:sSheetTitle.index('P')+3]}'"
            sSelectStringRows = f"Rpt_sheet = '{sSheetTitle}'"

            # Get worksheet
            ws = wb.create_sheet(title=sSheetTitle)


            self.write_titles_and_column_headers(ws, self.dtTitles, sSelectWorktable, self.sTitle_RR_YEAR, sTitle_WORKTABLE, iColumnCount, sSheetTitle)
            self.WriteFirst3ColumnsAndPageLayout(ws, self.dtLineSourceText, self.dtFootnotes, sSheetTitle, sSelectStringRows, sSelectWorktable)  
            # Freeze panes
            ws.freeze_panes = ws['D8']

            # Write hard values from dtaValue for all lines but 158
            col_map = {
                        iCurrentYear: 5,
                        iCurrentYear - 1: 7,
                        iCurrentYear - 2: 9,
                        iCurrentYear - 3: 11,
                        iCurrentYear - 4: 13
                    }
            for _, draValues in dtaValue[dtaValue["rpt_sheet"] == sSheetTitle].iterrows():
                if draValues["aline"] != 158:
                    iProcessYear = int(draValues["year"])
                    iROW_COUNT = int(draValues["aline"]) - iLineNumberOffset
                    if iProcessYear in col_map:
                        col = col_map[iProcessYear]
                        # Calculate the 'C' number for the named range, similar to the original VB
                        c_num = iCurrentYear - iProcessYear + 1
                        ws.cell(row=iROW_COUNT, column=col, value=draValues["value"])
                        cell = ws.cell(row=iROW_COUNT, column=col)
                        sNamedRange = f"A1L{draValues['aline']}C{c_num}"
                        cell.alignment = Alignment(horizontal="right")
                        wb.defined_names[sNamedRange] = DefinedName(name=sNamedRange, attr_text=f"'{sSheetTitle}'!${cell.column_letter}${cell.row}")
                        
                        cell.number_format = "#,##0"

            # Write hard values from dtaValue0_RR for line 158
            for _, draValues in dtaValue0_RR[dtaValue0_RR["rpt_sheet"] == sSheetTitle].iterrows():
                if draValues["aline"] == 158:
                    iProcessYear = int(draValues["year"])
                    iROW_COUNT = int(draValues["aline"]) - iLineNumberOffset

                    if iProcessYear in col_map:
                        c_num = iCurrentYear - iProcessYear + 1
                        col = col_map[iProcessYear]
                        value = draValues["value"] if iProcessYear == iCurrentYear else "0"
                        ws.cell(row=iROW_COUNT, column=col, value=value)
                        cell = ws.cell(row=iROW_COUNT, column=col)
                        sNamedRange = f"A1L{draValues['aline']}C{c_num}"
                        cell.alignment = Alignment(horizontal="right")
                        wb.defined_names[sNamedRange] = DefinedName(name=sNamedRange, attr_text=f"'{sSheetTitle}'!${cell.column_letter}${cell.row}")
                        
                        cell.number_format = "#,##0"

            # Write sources and values
            for _, drSource in dtLineSourceText[dtLineSourceText["rpt_sheet"] == sSheetTitle].iterrows():
                iLine = int(drSource["line"]) - iLineNumberOffset
                
                # Sources first
                source_cols = {
                    4: "c1", 6: "c2", 8: "c3", 10: "c4", 12: "c5",
                    14: "c6", 16: "c7", 18: "c8", 20: "c9"
                }
                for col, c_name in source_cols.items():
                    source_text = self.scrub_year(str(drSource.get(c_name, "")), self.current_year)
                    if source_text.startswith('=') or source_text.startswith('+'):
                        self.add_formula_as_text(ws, iLine, col, source_text)
                    else:
                        ws.cell(row=iLine, column=col, value=source_text)

                # C6 - C9 Values
                value_cols_c6_c9 = {15: "c6", 17: "c7", 19: "c8",  21: "c9"}
                for col, c_name in value_cols_c6_c9.items():
                    value = "0" if drSource["line"] == 158 else str(drSource[c_name])
                    c_num = (col - 15) // 2 + 6
                    ws.cell(row=iLine, column=col, value=value)
                    cell = ws.cell(row=iLine, column=col)
                    cell.number_format = "#,##0"
                    named_range_name = f"A1L{drSource['line']}C{c_num}"
                    wb.defined_names[named_range_name] = DefinedName(name=named_range_name, attr_text=f"'{sSheetTitle}'!${cell.column_letter}${cell.row}")
                    

                # Random values not from Data Dictionary
                random_lines = {111, 122, 131, 133, 137, 144, 145, 146, 151, 154, 155}
                if int(drSource["line"]) in random_lines:
                    random_value_cols = {5: "c1", 7: "c2", 9: "c3", 11: "c4", 13: "c5"}
                    for col, c_name in random_value_cols.items():
                        c_num = (col - 5) // 2 + 1
                        ws.cell(row=iLine, column=col, value=drSource.get(c_name, ""))
                        cell = ws.cell(row=iLine, column=col)
                        cell.number_format = "#,##0"
                        named_range_name = f"A1L{drSource['line']}C{c_num}"
                        wb.defined_names[named_range_name] = DefinedName(name=named_range_name, attr_text=f"'{sSheetTitle}'!${cell.column_letter}${cell.row}")
                        

            print(f"{sSheetTitle} completed")

        except Exception as ex:
            print(f"Error in A1P1: {ex}")
            import traceback
            print(traceback.format_exc())

    def A1P2A_worksheet(self, wb):
        try:
            sTitle_WORKTABLE = "WORKTABLE A1 PART 2A"
            iColumnCount = 57
            sSheetTitle = "A1P2A"
            iLineNumberOffset = 193

            dtaValue = self.dtAValue
            dtLineSourceText = self.dtLineSourceText
            iCurrentYear = int(self.current_year)

            print(f"Processing {sSheetTitle}")

            # Select worktable and string rows
            try:
                part_str = sSheetTitle[sSheetTitle.index('P')+1:sSheetTitle.index('P')+3]
            except Exception:
                part_str = sSheetTitle[sSheetTitle.index('P')+1:sSheetTitle.index('P')+2]
            sSelectWorktable = f"Worktable = '{sSheetTitle[:2]}' And Part = '{part_str}'"
            sSelectStringRows = f"Rpt_sheet = '{sSheetTitle}'"

            ws = wb.create_sheet(title=sSheetTitle)

            # Write titles and column headers
            self.write_titles_and_column_headers(ws, self.dtTitles, sSelectWorktable, self.sTitle_RR_YEAR, sTitle_WORKTABLE, iColumnCount, sSheetTitle)
            self.WriteFirst3ColumnsAndPageLayout(ws, self.dtLineSourceText, self.dtFootnotes, sSheetTitle, sSelectStringRows, sSelectWorktable)

            ws.freeze_panes = ws['D8']

            # Write hard values from dtaValue
            for _, draValues in dtaValue[dtaValue["rpt_sheet"] == sSheetTitle].iterrows():
                iProcessYear = int(draValues["year"])
                iROW_COUNT = int(draValues["aline"]) - iLineNumberOffset
                aCode_id = int(draValues["acode_id"]) if "acode_id" in draValues else int(draValues["aCode_id"]) if "aCode_id" in draValues else 0

                # Even/odd column mapping
                if aCode_id % 2 == 0:
                    col_map = {
                        iCurrentYear: 5,
                        iCurrentYear - 1: 11,
                        iCurrentYear - 2: 17,
                        iCurrentYear - 3: 23,
                        iCurrentYear - 4: 29
                    }
                    c_map = {
                        iCurrentYear: 1,
                        iCurrentYear - 1: 4,
                        iCurrentYear - 2: 7,
                        iCurrentYear - 3: 10,
                        iCurrentYear - 4: 13
                    }
                else:
                    col_map = {
                        iCurrentYear: 7,
                        iCurrentYear - 1: 13,
                        iCurrentYear - 2: 19,
                        iCurrentYear - 3: 25,
                        iCurrentYear - 4: 31
                    }
                    c_map = {
                        iCurrentYear: 2,
                        iCurrentYear - 1: 5,
                        iCurrentYear - 2: 8,
                        iCurrentYear - 3: 11,
                        iCurrentYear - 4: 14
                    }
                if iProcessYear in col_map:
                    col = col_map[iProcessYear]
                    c_num = c_map[iProcessYear]
                    cell = ws.cell(row=iROW_COUNT, column=col, value=draValues["value"])
                    cell.alignment = Alignment(horizontal="right")
                    cell.number_format = "#,##0"
                    sNamedRange = f"A1L{draValues['aline']}C{c_num}"
                    wb.defined_names[sNamedRange] = DefinedName(name=sNamedRange, attr_text=f"'{sSheetTitle}'!${cell.column_letter}${cell.row}")

            # Write sources and derived values
            for _, drSource in dtLineSourceText[dtLineSourceText["rpt_sheet"] == sSheetTitle].iterrows():
                iLine = int(drSource["line"]) - iLineNumberOffset

                # Write all 27 source columns (C1 to C27) to columns 4,6,8,...,56
                for n in range(1, 28):
                    col = 2 * n + 2  # 4,6,8,...,56
                    c_name = f"c{n}"
                    source_text = self.scrub_year(str(drSource.get(c_name, "")), iCurrentYear)
                    ws.cell(row=iLine, column=col, value=source_text if not (source_text.startswith('=') or source_text.startswith('+')) else "'" + source_text)

                # Write derived values to columns 9, 15, 21, ..., 57 (every 6th column starting from 9)
                derived_cols = [
                    (9, 3), (15, 6), (21, 9), (27, 12), (33, 15), (35, 16), (37, 17), (39, 18),
                    (41, 19), (43, 20), (45, 21), (47, 22), (49, 23), (51, 24), (53, 25), (55, 26), (57, 27)
                ]
                for col, c_num in derived_cols:
                    c_name = f"c{c_num}"
                    value = drSource.get(c_name, "")
                    cell = ws.cell(row=iLine, column=col, value=value)
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right")
                    named_range_name = f"A1L{drSource['line']}C{c_num}"
                    wb.defined_names[named_range_name] = DefinedName(name=named_range_name, attr_text=f"'{sSheetTitle}'!${cell.column_letter}${cell.row}")

            print(f"{sSheetTitle} completed")

        except Exception as ex:
            print(f"Error in {sSheetTitle}: {ex}")
            import traceback
            print(traceback.format_exc())

if __name__ == "__main__":
    
    create_reports = CreateReports(2023)
    create_reports.create_reports()

