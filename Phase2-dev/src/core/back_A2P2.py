def A2P2_worksheet_v1(self, wb):
        try:
            # BUILD THE WORKTABLE TITLE AND SET WORKTABLE VARS
            sTitle_WORKTABLE = "WORKTABLE A2 PART 2"
            iColumnCount = 91
            iWorkTableColumnCount = 44
            sSheetTitle = "A2P2"
            iLineNumberOffset = 193
            sNamedRangePrefix = "A2L"

            print(f"Processing {sSheetTitle}")

            # GET OUR SHEET AND A RANGE
            ws = wb.create_sheet(title=sSheetTitle)

            # WRITE OUT TITLE AND COLUMN HEADERS
            try:
                part_str = sSheetTitle[sSheetTitle.find('P')+1:sSheetTitle.find('P')+3]
            except:
                part_str = sSheetTitle[sSheetTitle.find('P')+1:sSheetTitle.find('P')+2]
            sSelectWorktable = f"Worktable = '{sSheetTitle[:2]}' And Part = '{part_str}'"
            self.write_titles_and_column_headers(ws, self.dtTitles, sSelectWorktable, self.sTitle_RR_YEAR, sTitle_WORKTABLE, iColumnCount, sSheetTitle)

            # WRITE OUT THE FIRST 3 COLUMNS
            sSelectStringRows = f"Rpt_sheet = '{sSheetTitle}'"
            self.WriteFirst3ColumnsAndPageLayout(ws, self.dtLineSourceText, self.dtFootnotes, sSheetTitle, sSelectStringRows, sSelectWorktable)

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
            iCurrentYear = int(self.current_year)
            for _, draValues in self.dtAValue[self.dtAValue["rpt_sheet"] == sSheetTitle].iterrows():
                iProcessYear = int(draValues["year"])
                aLine = int(draValues["aline"])
                acode_id = int(draValues["acode_id"])
                value_str = to_str(draValues["value"])

                drAnnPeriod = self.dtDataDictionary[self.dtDataDictionary["line"] == f"A2L{aLine}"]
                iROW_COUNT = aLine - iLineNumberOffset

                # Set code offsets and row offsets based on line ranges
                if aLine < 205:
                    iCodeOffset = 1311
                    iRowOffset = 8
                elif 216 < aLine < 219:
                    iCodeOffset = 1380
                    iRowOffset = 24
                elif 219 < aLine < 224:
                    iCodeOffset = 1394
                    iRowOffset = 27
                elif 235 < aLine < 238:
                    iCodeOffset = 1463
                    iRowOffset = 43
                elif 238 < aLine < 247:
                    iCodeOffset = 1476
                    iRowOffset = 46
                elif 258 < aLine < 261:
                    iCodeOffset = 1569
                    iRowOffset = 66

                # Set AnnPeriod for current year (excluding lines 219 and 238)
                if aLine != 219 and aLine != 238 and iProcessYear == iCurrentYear and not drAnnPeriod.empty:
                    set_cell(iROW_COUNT, 5, drAnnPeriod.iloc[0]["annperiod"], f"A2L{aLine}C1")

                # Price Index helper
                def get_pi(index, year_col_name):
                    pi_row = self.dtPriceIndexes[self.dtPriceIndexes['index'] == index]
                    return 0 if value_str == 0 else pi_row.iloc[0][year_col_name] if not pi_row.empty else 0

                # Check if line is in valid ranges
                is_valid_line = (aLine < 205 or 
                                (216 < aLine < 219) or 
                                (219 < aLine < 224) or 
                                (235 < aLine < 238) or 
                                (238 < aLine < 247) or 
                                (258 < aLine < 261))

                if not is_valid_line:
                    continue

                # Condition 1: aCode_id = iCodeOffset + 6 * (iROW_COUNT - iRowOffset)
                if acode_id == iCodeOffset + 6 * (iROW_COUNT - iRowOffset):
                    if iProcessYear == iCurrentYear:
                        set_cell(iROW_COUNT, 9, get_pi(1, "current_year"), f"A2L{aLine}C3", "0.0000")
                        set_cell(iROW_COUNT, 7, value_str, f"A2L{aLine}C2", "#,##0")
                    elif iProcessYear == iCurrentYear - 1:
                        set_cell(iROW_COUNT, 29, get_pi(1, "current_year_minus_1"), f"A2L{aLine}C13", "0.0000")
                        set_cell(iROW_COUNT, 27, value_str, f"A2L{aLine}C12", "#,##0")
                    elif iProcessYear == iCurrentYear - 2:
                        set_cell(iROW_COUNT, 45, get_pi(1, "current_year_minus_2"), f"A2L{aLine}C21", "0.0000")
                        set_cell(iROW_COUNT, 43, value_str, f"A2L{aLine}C20", "#,##0")
                    elif iProcessYear == iCurrentYear - 3:
                        set_cell(iROW_COUNT, 61, get_pi(1, "current_year_minus_3"), f"A2L{aLine}C29", "0.0000")
                        set_cell(iROW_COUNT, 59, value_str, f"A2L{aLine}C28", "#,##0")
                    elif iProcessYear == iCurrentYear - 4:
                        set_cell(iROW_COUNT, 77, get_pi(1, "current_year_minus_4"), f"A2L{aLine}C37", "0.0000")
                        set_cell(iROW_COUNT, 75, value_str, f"A2L{aLine}C36", "#,##0")

                # Condition 2: aCode_id = iCodeOffset + 6 * (iROW_COUNT - iRowOffset) + 1
                elif acode_id == iCodeOffset + 6 * (iROW_COUNT - iRowOffset) + 1:
                    if iProcessYear == iCurrentYear:
                        set_cell(iROW_COUNT, 13, get_pi(2, "current_year"), f"A2L{aLine}C5", "0.0000")
                        set_cell(iROW_COUNT, 11, value_str, f"A2L{aLine}C4", "#,##0")
                    elif iProcessYear == iCurrentYear - 1:
                        set_cell(iROW_COUNT, 33, get_pi(2, "current_year_minus_1"), f"A2L{aLine}C15", "0.0000")
                        set_cell(iROW_COUNT, 31, value_str, f"A2L{aLine}C14", "#,##0")
                    elif iProcessYear == iCurrentYear - 2:
                        set_cell(iROW_COUNT, 49, get_pi(2, "current_year_minus_2"), f"A2L{aLine}C23", "0.0000")
                        set_cell(iROW_COUNT, 47, value_str, f"A2L{aLine}C22", "#,##0")
                    elif iProcessYear == iCurrentYear - 3:
                        set_cell(iROW_COUNT, 65, get_pi(2, "current_year_minus_3"), f"A2L{aLine}C31", "0.0000")
                        set_cell(iROW_COUNT, 63, value_str, f"A2L{aLine}C30", "#,##0")
                    elif iProcessYear == iCurrentYear - 4:
                        set_cell(iROW_COUNT, 81, get_pi(2, "current_year_minus_4"), f"A2L{aLine}C39", "0.0000")
                        set_cell(iROW_COUNT, 79, value_str, f"A2L{aLine}C38", "#,##0")

                # Condition 3: Complex condition with specific aCode_id values
                elif (acode_id == iCodeOffset + 6 * (iROW_COUNT - iRowOffset) + 2 or
                      acode_id in {1341, 1345, 1349, 1353, 1357, 1361, 1368, 1372, 1376, 1424, 1428, 1432, 1436, 1440, 1444, 1451, 1455, 1459, 1530, 1534, 1538, 1542, 1546, 1550, 1557, 1561, 1565}):
                    if iProcessYear == iCurrentYear:
                        set_cell(iROW_COUNT, 17, get_pi(3, "current_year"), f"A2L{aLine}C7", "0.0000")
                        set_cell(iROW_COUNT, 15, value_str, f"A2L{aLine}C6", "#,##0")
                    elif iProcessYear == iCurrentYear - 1:
                        set_cell(iROW_COUNT, 37, get_pi(3, "current_year_minus_1"), f"A2L{aLine}C17", "0.0000")
                        set_cell(iROW_COUNT, 35, value_str, f"A2L{aLine}C16", "#,##0")
                    elif iProcessYear == iCurrentYear - 2:
                        set_cell(iROW_COUNT, 53, get_pi(3, "current_year_minus_2"), f"A2L{aLine}C25", "0.0000")
                        set_cell(iROW_COUNT, 51, value_str, f"A2L{aLine}C24", "#,##0")
                    elif iProcessYear == iCurrentYear - 3:
                        set_cell(iROW_COUNT, 69, get_pi(3, "current_year_minus_3"), f"A2L{aLine}C33", "0.0000")
                        set_cell(iROW_COUNT, 67, value_str, f"A2L{aLine}C32", "#,##0")
                    elif iProcessYear == iCurrentYear - 4:
                        set_cell(iROW_COUNT, 85, get_pi(3, "current_year_minus_4"), f"A2L{aLine}C41", "0.0000")
                        set_cell(iROW_COUNT, 83, value_str, f"A2L{aLine}C40", "#,##0")

                # Condition 4: Complex condition with specific aCode_id values
                elif (acode_id == iCodeOffset + 6 * (iROW_COUNT - iRowOffset) + 3 or
                      acode_id in {1335, 1338, 1342, 1346, 1350, 1354, 1358, 1362, 1365, 1369, 1373, 1377, 1418, 1421, 1425, 1429, 1433, 1437, 1441, 1445, 1448, 1452, 1456, 1460, 1524, 1527, 1531, 1535, 1539, 1543, 1547, 1551, 1554, 1558, 1562, 1566}):
                    if iProcessYear == iCurrentYear:
                        set_cell(iROW_COUNT, 21, get_pi(4, "current_year"), f"A2L{aLine}C9", "0.0000")
                        set_cell(iROW_COUNT, 19, value_str, f"A2L{aLine}C8", "#,##0")
                    elif iProcessYear == iCurrentYear - 1:
                        set_cell(iROW_COUNT, 41, get_pi(4, "current_year_minus_1"), f"A2L{aLine}C19", "0.0000")
                        set_cell(iROW_COUNT, 39, value_str, f"A2L{aLine}C18", "#,##0")
                    elif iProcessYear == iCurrentYear - 2:
                        set_cell(iROW_COUNT, 57, get_pi(4, "current_year_minus_2"), f"A2L{aLine}C27", "0.0000")
                        set_cell(iROW_COUNT, 55, value_str, f"A2L{aLine}C26", "#,##0")
                    elif iProcessYear == iCurrentYear - 3:
                        set_cell(iROW_COUNT, 73, get_pi(4, "current_year_minus_3"), f"A2L{aLine}C35", "0.0000")
                        set_cell(iROW_COUNT, 71, value_str, f"A2L{aLine}C34", "#,##0")
                    elif iProcessYear == iCurrentYear - 4:
                        set_cell(iROW_COUNT, 89, get_pi(4, "current_year_minus_4"), f"A2L{aLine}C43", "0.0000")
                        set_cell(iROW_COUNT, 87, value_str, f"A2L{aLine}C42", "#,##0")

                # Condition 5: Complex condition with specific aCode_id values
                elif (acode_id == iCodeOffset + 6 * (iROW_COUNT - iRowOffset) + 4 or
                      acode_id in {1336, 1339, 1343, 1347, 1351, 1355, 1359, 1363, 1366, 1370, 1374, 1378, 1419, 1422, 1426, 1430, 1434, 1438, 1442, 1446, 1449, 1453, 1457, 1461, 1525, 1528, 1532, 1536, 1540, 1544, 1548, 1552, 1555, 1559, 1563, 1567}):
                    if iProcessYear == iCurrentYear:
                        set_cell(iROW_COUNT, 23, value_str, f"A2L{aLine}C10", "#,##0")

                # Condition 6: Complex condition with specific aCode_id values
                elif (acode_id == iCodeOffset + 6 * (iROW_COUNT - iRowOffset) + 5 or
                      acode_id in {1337, 1340, 1344, 1348, 1352, 1356, 1360, 1364, 1367, 1371, 1375, 1379, 1420, 1423, 1427, 1431, 1435, 1439, 1443, 1447, 1450, 1454, 1458, 1462, 1526, 1529, 1533, 1537, 1541, 1545, 1549, 1553, 1556, 1560, 1564, 1568}):
                    if iProcessYear == iCurrentYear:
                        set_cell(iROW_COUNT, 25, value_str, f"A2L{aLine}C11", "#,##0")

            # WRITE RANGE NAMES FOR EMPTY CELLS
            for i in range(8, 70):  # Rows 8 to 69
                for j in range(5, 92):  # Columns 5 to 91
                    if i not in {26, 45, 68, 69} or j == 5:
                        caption_cell = ws.cell(row=6, column=j)
                        cell = ws.cell(row=i, column=j)

                        if caption_cell.value is not None and cell.value is None:
                            cell.value = "=NULL_VALUE"
                            caption_text = str(caption_cell.value).replace("(", "").replace(")", "")
                            named_range_name = f"A2L{i + iLineNumberOffset}{caption_text}"
                            wb.defined_names[named_range_name] = DefinedName(name=named_range_name, attr_text=f"'{sSheetTitle}'!${cell.column_letter}${cell.row}")
                            cell.alignment = Alignment(horizontal="right")
                            cell.number_format = "#######0"

            # WRITE OUT THE SOURCES AND ANY VALUES THAT EXECUTE THE SOURCE
            for _, drSource in self.dtLineSourceText[self.dtLineSourceText["rpt_sheet"] == sSheetTitle].iterrows():
                iLine = int(drSource["line"]) - iLineNumberOffset
                aline_str = str(drSource["line"])

                # Sources first - write all 43 source columns
                source_cols = [4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 38, 40, 42, 44, 46, 48, 50, 52, 54, 56, 58, 60, 62, 64, 66, 68, 70, 72, 74, 76, 78, 80, 82, 84, 86, 88]
                for idx, col_num in enumerate(source_cols):
                    c_name = f"c{idx + 1}"
                    source_text = self.scrub_year(to_str(drSource.get(c_name, "")), iCurrentYear)
                    ws.cell(row=iLine, column=col_num, value=f"'{source_text}" if source_text.startswith(('=', '+')) else source_text)

                # Handle special lines (26, 45, 68, 69) - Derived Values
                if iLine in {26, 45, 68, 69}:
                    # Set derived values for these special lines
                    set_cell(iLine, 7, drSource["c2"], f"A2L{aline_str}C2", "#,##0")
                    set_cell(iLine, 9, "=NULL_VALUE", f"A2L{aline_str}C3")
                    set_cell(iLine, 11, drSource["c4"], f"A2L{aline_str}C4", "#,##0")
                    set_cell(iLine, 13, "=NULL_VALUE", f"A2L{aline_str}C5")
                    set_cell(iLine, 15, drSource["c6"], f"A2L{aline_str}C6", "#,##0")
                    set_cell(iLine, 17, "=NULL_VALUE", f"A2L{aline_str}C7")
                    set_cell(iLine, 19, drSource["c8"], f"A2L{aline_str}C8", "#,##0")
                    set_cell(iLine, 21, "=NULL_VALUE", f"A2L{aline_str}C9")
                    set_cell(iLine, 23, drSource["c10"], f"A2L{aline_str}C10", "#,##0")
                    set_cell(iLine, 25, drSource["c11"], f"A2L{aline_str}C11", "#,##0")
                    set_cell(iLine, 27, drSource["c12"], f"A2L{aline_str}C12", "#,##0")
                    set_cell(iLine, 29, "=NULL_VALUE", f"A2L{aline_str}C13")
                    set_cell(iLine, 31, drSource["c14"], f"A2L{aline_str}C14", "#,##0")
                    set_cell(iLine, 33, "=NULL_VALUE", f"A2L{aline_str}C15")
                    set_cell(iLine, 35, drSource["c16"], f"A2L{aline_str}C16", "#,##0")
                    set_cell(iLine, 37, "=NULL_VALUE", f"A2L{aline_str}C17")
                    set_cell(iLine, 39, drSource["c18"], f"A2L{aline_str}C18", "#,##0")
                    set_cell(iLine, 41, "=NULL_VALUE", f"A2L{aline_str}C19")
                    set_cell(iLine, 43, drSource["c20"], f"A2L{aline_str}C20", "#,##0")
                    set_cell(iLine, 45, "=NULL_VALUE", f"A2L{aline_str}C21")
                    set_cell(iLine, 47, drSource["c22"], f"A2L{aline_str}C22", "#,##0")
                    set_cell(iLine, 49, "=NULL_VALUE", f"A2L{aline_str}C23")
                    set_cell(iLine, 51, drSource["c24"], f"A2L{aline_str}C24", "#,##0")
                    set_cell(iLine, 53, "=NULL_VALUE", f"A2L{aline_str}C25")
                    set_cell(iLine, 55, drSource["c26"], f"A2L{aline_str}C26", "#,##0")
                    set_cell(iLine, 57, "=NULL_VALUE", f"A2L{aline_str}C27")
                    set_cell(iLine, 59, drSource["c28"], f"A2L{aline_str}C28", "#,##0")
                    set_cell(iLine, 61, "=NULL_VALUE", f"A2L{aline_str}C29")
                    set_cell(iLine, 63, drSource["c30"], f"A2L{aline_str}C30", "#,##0")
                    set_cell(iLine, 65, "=NULL_VALUE", f"A2L{aline_str}C31")
                    set_cell(iLine, 67, drSource["c32"], f"A2L{aline_str}C32", "#,##0")
                    set_cell(iLine, 69, "=NULL_VALUE", f"A2L{aline_str}C33")
                    set_cell(iLine, 71, drSource["c34"], f"A2L{aline_str}C34", "#,##0")
                    set_cell(iLine, 73, "=NULL_VALUE", f"A2L{aline_str}C35")
                    set_cell(iLine, 75, drSource["c36"], f"A2L{aline_str}C36", "#,##0")
                    set_cell(iLine, 77, "=NULL_VALUE", f"A2L{aline_str}C37")
                    set_cell(iLine, 79, drSource["c38"], f"A2L{aline_str}C38", "#,##0")
                    set_cell(iLine, 81, "=NULL_VALUE", f"A2L{aline_str}C39")
                    set_cell(iLine, 83, drSource["c40"], f"A2L{aline_str}C40", "#,##0")
                    set_cell(iLine, 85, "=NULL_VALUE", f"A2L{aline_str}C41")
                    set_cell(iLine, 87, drSource["c42"], f"A2L{aline_str}C42", "#,##0")
                    set_cell(iLine, 89, "=NULL_VALUE", f"A2L{aline_str}C43")
                    ws.cell(row=iLine, column=90, value=f"'{self.scrub_year(drSource.get('c44', ''), iCurrentYear)}")
                    set_cell(iLine, 91, drSource["c44"], f"A2L{aline_str}C44", "#,##0")
                else:
                    # Create source for C44 using GetSourceForA2SummaryColumn equivalent
                    sSource = self.get_source_for_a2_summary_column(ws, drSource["line"], iLine)
                    ws.cell(row=iLine, column=90, value=f"'{sSource if len(sSource) > 0 else ''}")
                    set_cell(iLine, 91, sSource, f"A2L{aline_str}C44", "#,##0")

            self.format_all_cells(ws)
            print(f"{sSheetTitle} completed")

        except Exception as ex:
            print(f"Error in {sSheetTitle}: {ex}")
            import traceback
            print(traceback.format_exc())
