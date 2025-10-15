from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from report_builder.core.registry import register
from report_builder.core.context import ReportContext

def to_str(value):
    """Helper function to convert value to string safely"""
    return str(value) if value is not None else ""

def apostrophe(text):
    """Add apostrophe prefix if text starts with = or +"""
    return f"'{text}" if text.startswith(('=', '+')) else text

@register("A3P2_worksheet")
def A3P2_worksheet(ctx: ReportContext, wb: Workbook):
    try:
        sTitle_WORKTABLE = "WORKTABLE A3 PART 2"
        iColumnCount = 89
        iWorkTableColumnCount = 43
        sSheetTitle = "A3P2"
        iLineNumberOffset = 193
        sNamedRangePrefix = "A3L"

        print(f"Processing {sSheetTitle}")
        ws = wb.create_sheet(title=sSheetTitle)
        ws.freeze_panes = ws['D8']

        part_str = sSheetTitle[sSheetTitle.find('P')+1:]
        part_only = part_str[:2] if len(part_str) >= 2 and part_str[:2].isdigit() else part_str[:1]
        sSelectWorktable = f"Worktable = '{sSheetTitle[:2]}' And Part = '{part_only}'"
        sSelectStringRows = f"Rpt_sheet = '{sSheetTitle}'"

        ctx.write_titles_and_column_headers(ws, ctx.dtTitles, sSelectWorktable,
                                            ctx.sTitle_RR_YEAR, sTitle_WORKTABLE,
                                            iColumnCount, sSheetTitle)
        ctx.WriteFirst3ColumnsAndPageLayout(ws, ctx.dtLineSourceText, ctx.dtFootnotes,
                                            sSheetTitle, sSelectStringRows, sSelectWorktable)

        def set_cell(row, col, value, name=None, num_format=None, align_right=True):
            c = ws.cell(row=row, column=col, value=value)
            if name:
                wb.defined_names[name] = DefinedName(
                    name=name, attr_text=f"'{sSheetTitle}'!${c.column_letter}${c.row}"
                )
            if align_right:
                c.alignment = Alignment(horizontal="right")
            if num_format:
                c.number_format = num_format
            return c

        def get_pi9(year_col_name):
            df = ctx.dtPriceIndexes[ctx.dtPriceIndexes['index'] == 9]
            return "" if df.empty else df.iloc[0][year_col_name]

        iCurrentYear = int(ctx.current_year)
        special_lines = {219, 224}

        # WRITE HARD VALUES
        for _, r in ctx.variable_ctx.dtAValue[ctx.variable_ctx.dtAValue["rpt_sheet"] == sSheetTitle].iterrows():
            aLine = int(r["aline"])
            if aLine in special_lines:
                continue

            iProcessYear = int(r["year"])
            iROW_COUNT = aLine - iLineNumberOffset
            aCol = to_str(r.get("acolumn", r.get("acol", r.get("aColumn", ""))))  # tolerant
            value = to_str(r["value"])

            if iProcessYear == iCurrentYear:
                # AnnPeriod (C1)
                dd = ctx.dtDataDictionary[ctx.dtDataDictionary["line"] == f"{sNamedRangePrefix}{aLine}"]
                if not dd.empty:
                    set_cell(iROW_COUNT, 5, dd.iloc[0]["annperiod"], f"{sNamedRangePrefix}{aLine}C1")
                # C2 price index (Index 9)
                set_cell(iROW_COUNT, 7, get_pi9("current_year"), f"{sNamedRangePrefix}{aLine}C2", "0.0000")
                # C3..C8 based on aColumn 3..8
                if aCol == "3":
                    set_cell(iROW_COUNT, 9,  value, f"{sNamedRangePrefix}{aLine}C3", "#,##0")
                if aCol == "4":
                    formula_prefix = "=A3L215C4+A3L216C4+" if iROW_COUNT == 25 else ""
                    set_cell(iROW_COUNT, 11, f"{formula_prefix}{value}", f"{sNamedRangePrefix}{aLine}C4", "#,##0")
                # Continue with remaining columns logic...

            # Continue with other years...

        # SOURCES
        for _, dr in ctx.dtLineSourceText[ctx.dtLineSourceText["rpt_sheet"] == sSheetTitle].iterrows():
            line = int(dr["line"])
            iLine = line - iLineNumberOffset

            # C1..C43 -> cols 4..88 (even)
            for idx, col in enumerate(range(4, 89, 2), start=1):
                key = f"c{idx}"
                ws.cell(row=iLine, column=col, value=apostrophe(ctx.scrub_year(to_str(dr.get(key, "")), iCurrentYear)))

            if line in special_lines:
                # Complex NULL/value pattern implementation
                pass  # Implement detailed pattern
            else:
                # Create sources for C37..C42 and C43
                for k, base_col in zip(range(3, 9), [76, 78, 80, 82, 84, 86]):
                    sSource = ctx.get_source_for_a3p2_summary_column(ws, line, iLine, k)
                    ws.cell(row=iLine, column=base_col, value=apostrophe(sSource) if len(sSource) > 0 else "")
                    set_cell(iLine, base_col + 1, sSource, f"{sNamedRangePrefix}{line}C{36 + k}", "#,##0")

                set_cell(iLine, 89, ctx.scrub_year(to_str(dr.get("c43", "")), iCurrentYear),
                        f"{sNamedRangePrefix}{line}C43", "#,##0")

        ctx.format_all_cells(ws)
        print(f"{sSheetTitle} completed")
    except Exception as ex:
        print(f"Error in {sSheetTitle}: {ex}")
        import traceback
        print(traceback.format_exc())
