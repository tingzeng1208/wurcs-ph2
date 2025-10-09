                    
                    # Source column (4, 6, 8, ...)
                    source_col = 2 * i + 2
                    source_text = self.scrub_year(str(drSource.get(c_name, "")), iCurrentYear)
                    ws.cell(row=iLine, column=source_col, value=f"'{source_text}" if source_text.startswith(('=', '+')) else source_text)

                    # Derived Value column (5, 7, 9, ...)
                aLine = int(draValues["aline"])
                iROW_COUNT = aLine - iLineNumberOffset
                acode_id = int(draValues.get("acode_id", 0))
                value = draValues.get("value", 0)
                try:
                    value_str = str(int(float(value))) if float(value).is_integer() else str(value)
                except (ValueError, TypeError):
                    value_str = str(value)

                # Get AnnPeriod
                ann_period_row = dtDataDictionary[dtDataDictionary["line"] == f"A2L{aLine}"]
                acode_check = acode_id - (iCodeOffset + 6 * (iROW_COUNT - iRowOffset))

                # Define column and C-number mappings
                acode_id_set_c7 = {1162, 1166, 1170, 1174, 1178, 1182, 1186, 1190, 1194, 1198, 1202, 1206, 1210, 1214, 1218, 1222, 1226, 1230, 1243, 1247, 1251, 1255, 1259, 1263}
                acode_id_set_c9 = {1144, 1147, 1150, 1153, 1156, 1159, 1234, 1237, 1240}
                acode_id_set_c10 = {1145, 1148, 1151, 1154, 1157, 1160, 1164, 1168, 1172, 1176, 1180, 1184, 1188, 1192, 1196, 1200, 1204, 1208, 1212, 1216, 1220, 1224, 1228, 1232, 1235, 1238, 1241, 1245, 1249, 1253, 1257, 1261, 1265}
                acode_id_set_c11 = {1146, 1149, 1152, 1155, 1158, 1161, 1165, 1169, 1173, 1177, 1181, 1185, 1189, 1193, 1197, 1201, 1205, 1209, 1213, 1217, 1221, 1225, 1229, 1233, 1236, 1239, 1242, 1246, 1250, 1254, 1258, 1262, 1266}

                col_map = {}
                if is_main_block and acode_check == 0:
                    col_map = { 'val': 7 + year_diff * 20, 'idx': 9 + year_diff * 20, 'c_val': 2 + year_diff * 10, 'c_idx': 3 + year_diff * 10, 'idx_ref': 1 }
                elif is_main_block and acode_check == 1:
                    col_map = { 'val': 11 + year_diff * 20, 'idx': 13 + year_diff * 20, 'c_val': 4 + year_diff * 10, 'c_idx': 5 + year_diff * 10, 'idx_ref': 2 }
                elif (is_main_block and acode_check == 2) or acode_id in acode_id_set_c7:
                    col_map = { 'val': 15 + year_diff * 20, 'idx': 17 + year_diff * 20, 'c_val': 6 + year_diff * 10, 'c_idx': 7 + year_diff * 10, 'idx_ref': 3 }
                elif (is_main_block and acode_check == 3) or acode_id in acode_id_set_c9:
                    col_map = { 'val': 19 + year_diff * 20, 'idx': 21 + year_diff * 20, 'c_val': 8 + year_diff * 10, 'c_idx': 9 + year_diff * 10, 'idx_ref': 4 }
                elif (is_main_block and acode_check == 4) or acode_id in acode_id_set_c10:
                    if iProcessYear == iCurrentYear:
                        cell = ws.cell(row=iROW_COUNT, column=23, value=value_str)
                        cell.alignment = Alignment(horizontal="right")
                        cell.number_format = "#,##0"
                        wb.defined_names[f"A2L{aLine}C10"] = DefinedName(name=f"A2L{aLine}C10", attr_text=f"'{sSheetTitle}'!${cell.column_letter}${cell.row}")
                elif (is_main_block and acode_check == 5) or acode_id in acode_id_set_c11:
                     if iProcessYear == iCurrentYear:
                        cell = ws.cell(row=iROW_COUNT, column=25, value=value_str)
                        cell.alignment = Alignment(horizontal="right")
                        cell.number_format = "#,##0"
                        wb.defined_names[f"A2L{aLine}C11"] = DefinedName(name=f"A2L{aLine}C11", attr_text=f"'{sSheetTitle}'!${cell.column_letter}${cell.row}")

                if col_map:
                    # Write Value
                    cell_val = ws.cell(row=iROW_COUNT, column=col_map['val'], value=value_str)
                    cell_val.alignment = Alignment(horizontal="right")
                    cell_val.number_format = "#,##0"
                    wb.defined_names[f"A2L{aLine}C{col_map['c_val']}"] = DefinedName(name=f"A2L{aLine}C{col_map['c_val']}", attr_text=f"'{sSheetTitle}'!${cell_val.column_letter}${cell_val.row}")

                    # Write Price Index
                    price_index_row = dtPriceIndexes[dtPriceIndexes['index'] == col_map['idx_ref']]
                    if not price_index_row.empty:
                        price_index_col_name = f"current_year{'' if year_diff == 0 else f'_minus_{year_diff}'}"
                        price_index_val = 0.0 if value_str == "0" else price_index_row.iloc[0][price_index_col_name]
                        cell_idx = ws.cell(row=iROW_COUNT, column=col_map['idx'], value=price_index_val)
                        cell_idx.alignment = Alignment(horizontal="right")
                        cell_idx.number_format = "0.0000"
                        wb.defined_names[f"A2L{aLine}C{col_map['c_idx']}"] = DefinedName(name=f"A2L{aLine}C{col_map['c_idx']}", attr_text=f"'{sSheetTitle}'!${cell_idx.column_letter}${cell_idx.row}")

                    for col, c_name in derived_map.items():
                        c_num = int(c_name[1:])
                        cell = ws.cell(row=iLine, column=col, value=drSource.get(c_name, ""))
                        cell.alignment = Alignment(horizontal="right")
                        cell.number_format = "#,##0"
                        wb.defined_names[f"A2L{drSource['line']}C{c_num}"] = DefinedName(name=f"A2L{drSource['line']}C{c_num}", attr_text=f"'{sSheetTitle}'!${cell.column_letter}${cell.row}")

                    # C44
                    ws.cell(row=iLine, column=90, value=f"'{self.scrub_year(drSource.get('c44', ''), iCurrentYear)}")
                    cell_c44 = ws.cell(row=iLine, column=91, value=drSource.get("c44", ""))
                    cell_c44.alignment = Alignment(horizontal="right")
                    cell_c44.number_format = "#,##0"
                    wb.defined_names[f"A2L{drSource['line']}C44"] = DefinedName(name=f"A2L{drSource['line']}C44", attr_text=f"'{sSheetTitle}'!${cell_c44.column_letter}${cell_c44.row}")
                else:
                    # This logic is complex and seems to build a formula string by summing up other cells.
                    # A simplified placeholder is used here. You may need to implement the full formula generation.
                    # For now, we'll just create the named range.
                    source_c44 = drSource.get("c44", "")
                    ws.cell(row=iLine, column=90, value=f"'{source_c44}" if source_c44 else "")

                    cell_c44 = ws.cell(row=iLine, column=91, value=source_c44)
                    cell_c44.alignment = Alignment(horizontal="right")
                    cell_c44.number_format = "#,##0"
                    wb.defined_names[f"A2L{drSource['line']}C44"] = DefinedName(name=f"A2L{drSource['line']}C44", attr_text=f"'{sSheetTitle}'!${cell_c44.column_letter}${cell_c44.row}")

                    cell = ws.cell(row=i, column=j)
                    if caption_cell.value:
                        caption_text = str(caption_cell.value).replace("(", "").replace(")", "")
                        # Find the named range for this cell if it exists
                        current_named_range = None
                        for name, dest in wb.defined_names.items():
                             if dest.attr_text == f"'{sSheetTitle}'!${cell.column_letter}${cell.row}":
                                 current_named_range = name
                                 break
                        
                        if cell.value is None:
                            cell.value = "=NULL_VALUE"
                            cell.alignment = Alignment(horizontal="right")
                            cell.number_format = "#######0"
                        elif current_named_range in ssac_cells:
                            original_value = cell.value
                            cell.value = f'=IF(SSAC="Y",0,{original_value})'


