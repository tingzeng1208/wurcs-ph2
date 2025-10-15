from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from report_builder.core.registry import register
from report_builder.core.context import ReportContext

@register("A2P1_worksheet")
def A2P1_worksheet(ctx: ReportContext, wb: Workbook):
    try:
        # BUILD THE WORKTABLE TITLE AND SET WORKTABLE VARS
        sTitle_WORKTABLE = "WORKTABLE A2 PART 1"
        iColumnCount = 91
        iWorkTableColumnCount = 44
        sSheetTitle = "A2P1"
        iLineNumberOffset = 93
        sNamedRangePrefix = "A2L"

        print(f"Processing {sSheetTitle}")

        # GET OUR SHEET AND A RANGE
        ws = wb.create_sheet(title=sSheetTitle)

        # WRITE OUT TITLE AND COLUMN HEADERS
        part_str = sSheetTitle[sSheetTitle.find('P')+1:]
        sSelectWorktable = f"Worktable = '{sSheetTitle[:2]}' And Part = '{part_str}'"
        write_titles_and_column_headers(ctx, ws, ctx.dtTitles, sSelectWorktable, ctx.sTitle_RR_YEAR, sTitle_WORKTABLE, iColumnCount, sSheetTitle)

        # WRITE OUT THE FIRST 3 COLUMNS
        sSelectStringRows = f"Rpt_sheet = '{sSheetTitle}'"
        WriteFirst3ColumnsAndPageLayout(ctx, ws, ctx.dtLineSourceText, ctx.dtFootnotes, sSheetTitle, sSelectStringRows, sSelectWorktable)

        # FREEZE THE ROWS AND COLUMNS
        ws.freeze_panes = ws['D8']

        # Helper to set cell value and named range
        def set_cell(row, col, value, name, num_format=None):
            cell = ws.cell(row=row, column=col, value=value)
            wb.defined_names[name] = DefinedName(name=name, attr_text=f"'{sSheetTitle}'!${cell.column_letter}${cell.row}")
            cell.alignment = Alignment(horizontal="right")
            if num_format:
                cell.number_format = num_format
            return cell

        # WRITE HARD VALUES FROM AVALUES
        iCurrentYear = int(ctx.current_year)
        dtAValue = ctx.variable_ctx.dtAValue
        dtDataDictionary = ctx.dtDataDictionary
        dtPriceIndexes = ctx.variable_ctx.dtPriceIndexes
        
        for _, draValues in dtAValue[dtAValue["rpt_sheet"] == sSheetTitle].iterrows():
            iProcessYear = int(draValues["year"])
            aLine = int(draValues["aline"])
            acode_id = int(draValues["acode_id"])
            value_str = to_str(draValues["value"])
            
            iROW_COUNT = aLine - iLineNumberOffset
            drAnnPeriod = dtDataDictionary[dtDataDictionary["line"] == f"A2L{aLine}"]
            
            # Set code offsets and row offsets based on line number
            if aLine < 142:
                iCodeOffset = 898
                iRowOffset = 8
            else:
                iCodeOffset = 1267
                iRowOffset = 82

            # Set AnnPeriod for current year
            if iProcessYear == iCurrentYear and not drAnnPeriod.empty:
                set_cell(iROW_COUNT, 5, str(drAnnPeriod.iloc[0]["annperiod"]), f"A2L{aLine}C1")

            # Process different aCode_id conditions
            line_conditions = (aLine < 142 or (aLine > 174 and aLine < 181))
            
            if line_conditions and acode_id == iCodeOffset + 6 * (iROW_COUNT - iRowOffset):
                # Index 1 processing
                if iProcessYear == iCurrentYear:
                    price_index_val = "0" if value_str == "0" else str(dtPriceIndexes[dtPriceIndexes["index"] == 1].iloc[0]["current_year"])
                    set_cell(iROW_COUNT, 9, price_index_val, f"A2L{aLine}C3", "0.0000")
                    set_cell(iROW_COUNT, 7, value_str, f"A2L{aLine}C2")
                    
                elif iProcessYear == iCurrentYear - 1:
                    price_index_val = "0" if value_str == "0" else str(dtPriceIndexes[dtPriceIndexes["index"] == 1].iloc[0]["current_year_minus_1"])
                    set_cell(iROW_COUNT, 29, price_index_val, f"A2L{aLine}C13", "0.0000")
                    set_cell(iROW_COUNT, 27, value_str, f"A2L{aLine}C12")
                    
                # Continue for other years...

            elif line_conditions and acode_id == iCodeOffset + 6 * (iROW_COUNT - iRowOffset) + 1:
                # Index 2 processing
                if iProcessYear == iCurrentYear:
                    price_index_val = "0" if value_str == "0" else str(dtPriceIndexes[dtPriceIndexes["index"] == 2].iloc[0]["current_year"])
                    set_cell(iROW_COUNT, 13, price_index_val, f"A2L{aLine}C5", "0.0000")
                    set_cell(iROW_COUNT, 11, value_str, f"A2L{aLine}C4")
                    
                # Continue for other years...

        # Write sources and derived values
        for _, drSource in ctx.dtLineSourceText[ctx.dtLineSourceText["rpt_sheet"] == sSheetTitle].iterrows():
            iLine = int(drSource["line"]) - iLineNumberOffset

            # Write all source columns
            for n in range(1, 44):
                col = 2 * n + 2  # 4, 6, 8, ..., 88
                c_name = f"c{n}"
                source_text = scrub_year(str(drSource.get(c_name, "")), iCurrentYear)
                ws.cell(row=iLine, column=col, value=f"'{source_text}" if not source_text.startswith(('=', '+')) else source_text)

            # Handle special line (88) with derived values
            if iLine == 88:
                # Set derived values for line 88
                derived_values = [
                    (7, 2), (9, 3), (11, 4), (13, 5), (15, 6), (17, 7), (19, 8), (21, 9),
                    (23, 10), (25, 11), (27, 12), (29, 13), (31, 14), (33, 15), (35, 16),
                    (37, 17), (39, 18), (41, 19), (43, 20), (45, 21), (47, 22), (49, 23),
                    (51, 24), (53, 25), (55, 26), (57, 27), (59, 28), (61, 29), (63, 30),
                    (65, 31), (67, 32), (69, 33), (71, 34), (73, 35), (75, 36), (77, 37),
                    (79, 38), (81, 39), (83, 40), (85, 41), (87, 42), (89, 43), (91, 44)
                ]
                
                for col, c_num in derived_values:
                    c_name = f"c{c_num}"
                    if c_num in [3, 5, 7, 9, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35, 37, 39, 41, 43]:
                        # NULL_VALUE columns
                        cell = ws.cell(row=iLine, column=col, value="=NULL_VALUE")
                        cell.alignment = Alignment(horizontal="right")
                    else:
                        # Regular derived values
                        cell = ws.cell(row=iLine, column=col, value=str(drSource.get(c_name, "")))
                        cell.number_format = "#,##0"
                    
                    named_range = f"A2L{drSource['line']}C{c_num}"
                    wb.defined_names[named_range] = DefinedName(name=named_range, attr_text=f"'{sSheetTitle}'!${cell.column_letter}${cell.row}")
            else:
                # Create source for C44 using summary column function
                sSource = get_source_for_a2_summary_column(ctx, str(drSource["line"]), iLine)
                ws.cell(row=iLine, column=90, value=f"'{sSource}" if len(sSource) > 0 else "")
                cell = ws.cell(row=iLine, column=91, value=sSource)
                named_range = f"A2L{drSource['line']}C44"
                wb.defined_names[named_range] = DefinedName(name=named_range, attr_text=f"'{sSheetTitle}'!${cell.column_letter}${cell.row}")

        print(f"{sSheetTitle} completed")

    except Exception as ex:
        print(f"Error in {sSheetTitle}: {ex}")
        import traceback
        print(traceback.format_exc())

# Helper functions
def write_titles_and_column_headers(ctx, ws, dtTitles, sSelectWorktable, sTitle_RR_YEAR, sTitle_WORKTABLE, iColumnCount, sSheetTitle):
    # Implementation would go here
    pass

def WriteFirst3ColumnsAndPageLayout(ctx, ws, dtLineSourceText, dtFootnotes, sSheetTitle, sSelectStringRows, sSelectWorktable):
    # Implementation would go here
    pass

def scrub_year(text, current_year):
    # Implementation would go here
    return text

def to_str(value):
    return str(value) if value is not None else ""

def get_source_for_a2_summary_column(ctx, line_number, cell_line_number):
    # Implementation would go here (from original file)
    return ""
