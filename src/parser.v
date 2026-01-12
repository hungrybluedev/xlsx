module xlsx

import os
import compress.szip
import rand
import encoding.xml

fn create_temporary_directory() string {
	for {
		location := os.join_path(os.temp_dir(), 'xlsx-${rand.hex(10)}')
		if os.exists(location) {
			continue
		}
		os.mkdir(location) or { continue }
		return location
	}
	// Should not reach here
	return ''
}

fn load_shared_strings(path string, shared_strings_path string) ![]string {
	mut shared_strings := []string{}

	if !os.exists(shared_strings_path) {
		return shared_strings
	}

	strings_doc := xml.XMLDocument.from_file(shared_strings_path) or {
		return error('Failed to parse shared strings file of excel file: ${path}')
	}

	all_defined_strings := strings_doc.get_elements_by_tag('si')
	for definition in all_defined_strings {
		if definition.children.len == 0 {
			return error('Invalid shared string definition: empty <si> element')
		}
		t_element := definition.children[0]
		if t_element !is xml.XMLNode || (t_element as xml.XMLNode).name != 't' {
			return error('Invalid shared string definition: expected <t> element, found: ${definition}')
		}

		t_node := t_element as xml.XMLNode
		if t_node.children.len == 0 {
			// Empty string - this is valid
			shared_strings << ''
			continue
		}
		content := t_node.children[0]
		if content !is string {
			return error('Invalid shared string definition: expected text content, found: ${definition}')
		}
		shared_strings << (content as string)
	}

	return shared_strings
}

fn load_worksheets_metadata(path string, worksheets_file_path string) !map[int]string {
	if !os.exists(worksheets_file_path) {
		return error('Worksheets file does not exist: ${path}')
	}
	worksheets_doc := xml.XMLDocument.from_file(worksheets_file_path) or {
		return error('Failed to parse worksheets file of excel file: ${path}')
	}

	worksheets := worksheets_doc.get_elements_by_tag('sheet')
	mut worksheets_metadata := map[int]string{}

	for worksheet in worksheets {
		worksheets_metadata[worksheet.attributes['sheetId'].int()] = worksheet.attributes['name']
	}
	return worksheets_metadata
}

fn load_styles(path string, styles_path string) !StyleInfo {
	mut style_info := StyleInfo{
		num_fmt_codes: map[int]string{}
		cell_xfs:      []CellXf{}
		fills:         []FillInfo{}
	}

	if !os.exists(styles_path) {
		return style_info
	}

	styles_doc := xml.XMLDocument.from_file(styles_path) or {
		return error('Failed to parse styles file of excel file: ${path}')
	}

	// Parse numFmts section
	num_fmts := styles_doc.get_elements_by_tag('numFmt')
	for num_fmt in num_fmts {
		num_fmt_id := (num_fmt.attributes['numFmtId'] or { '0' }).int()
		format_code := num_fmt.attributes['formatCode'] or { '' }
		style_info.num_fmt_codes[num_fmt_id] = format_code
	}

	// Parse fills section
	fills := styles_doc.get_elements_by_tag('fill')
	for fill in fills {
		mut fill_info := FillInfo{
			pattern_type: 'none'
			theme:        none
			tint:         none
		}

		// Look for patternFill element
		for child in fill.children {
			if child is xml.XMLNode && child.name == 'patternFill' {
				pattern_fill := child as xml.XMLNode
				fill_info.pattern_type = pattern_fill.attributes['patternType'] or { 'none' }

				// Look for fgColor element with theme
				for fg_child in pattern_fill.children {
					if fg_child is xml.XMLNode && fg_child.name == 'fgColor' {
						fg_color := fg_child as xml.XMLNode
						if theme_str := fg_color.attributes['theme'] {
							fill_info.theme = theme_str.int()
						}
						if tint_str := fg_color.attributes['tint'] {
							fill_info.tint = tint_str.f64()
						}
					}
				}
			}
		}
		style_info.fills << fill_info
	}

	// Parse cellXfs section
	// Note: get_elements_by_tag returns ALL xf elements from both cellStyleXfs and cellXfs
	// cellStyleXfs entries do NOT have xfId attribute
	// cellXfs entries HAVE xfId attribute (referencing cellStyleXfs)
	xfs := styles_doc.get_elements_by_tag('xf')
	for xf in xfs {
		// Only include xf entries that have xfId attribute (these are cellXfs entries)
		// Skip cellStyleXfs entries (which don't have xfId)
		if _ := xf.attributes['xfId'] {
			num_fmt_id := (xf.attributes['numFmtId'] or { '0' }).int()
			fill_id := (xf.attributes['fillId'] or { '0' }).int()
			style_info.cell_xfs << CellXf{
				num_fmt_id: num_fmt_id
				fill_id:    fill_id
			}
		}
	}

	return style_info
}

// Detect currency from Excel format code
fn detect_currency_from_format(format_code string) ?Currency {
	// Check for currency symbols in format codes
	// Excel uses locale IDs like [$£-809] for GBP, [$$-409] for USD, etc.
	if format_code.contains('£') || format_code.contains('-809]') {
		return .gbp
	}
	if format_code.contains('[$$') || format_code.contains('-409]') {
		return .usd
	}
	if format_code.contains('€') || format_code.contains('-407]') {
		return .eur
	}
	if format_code.contains('¥') {
		// Check locale to distinguish JPY from CNY
		if format_code.contains('-411]') {
			return .jpy
		}
		if format_code.contains('-804]') {
			return .cny
		}
		// Default to JPY if just yen symbol without locale
		return .jpy
	}
	if format_code.contains('₹') || format_code.contains('-4009]') {
		return .inr
	}
	return none
}

pub fn Document.from_file(path string) !Document {
	// Fail if the file does not exist.
	if !os.exists(path) {
		return error('File does not exist: ${path}')
	}
	// First, we extract the ZIP file into a temporary directory.
	location := create_temporary_directory()

	szip.extract_zip_to_dir(path, location) or {
		return error('Failed to extract information from file: ${path}\nError:\n${err}')
	}

	// Then we list the files in the "xl" directory.
	xl_path := os.join_path(location, 'xl')

	// Load the strings from the shared strings file, if it exists.
	shared_strings_path := os.join_path(xl_path, 'sharedStrings.xml')
	shared_strings := load_shared_strings(path, shared_strings_path)!

	// Load the styles from the styles file, if it exists.
	styles_path := os.join_path(xl_path, 'styles.xml')
	style_info := load_styles(path, styles_path)!

	// Load the sheets metadata from the workbook file.
	worksheets_file_path := os.join_path(xl_path, 'workbook.xml')
	sheet_metadata := load_worksheets_metadata(path, worksheets_file_path)!

	// Finally, we can load the sheets.
	all_sheet_paths := os.ls(os.join_path(xl_path, 'worksheets'))!

	mut sheet_map := map[int]Sheet{}

	for sheet_file in all_sheet_paths {
		sheet_path := os.join_path(xl_path, 'worksheets', sheet_file)
		sheet_id := sheet_file.all_after('sheet').all_before('.xml').int()
		sheet_name := sheet_metadata[sheet_id] or {
			return error('Failed to find sheet name for sheet ID: ${sheet_id}')
		}

		sheet_doc := xml.XMLDocument.from_file(sheet_path) or {
			return error('Failed to parse sheet file: ${sheet_path}')
		}

		sheet := Sheet.from_doc(sheet_name, sheet_doc, shared_strings, style_info) or {
			return error('Failed to parse sheet file: ${sheet_path}')
		}

		sheet_map[sheet_id] = sheet
	}

	return Document{
		shared_strings: shared_strings
		sheets:         sheet_map
	}
}

fn Sheet.from_doc(name string, doc xml.XMLDocument, shared_strings []string, style_info StyleInfo) !Sheet {
	dimension_tags := doc.get_elements_by_tag('dimension')
	if dimension_tags.len != 1 {
		return error('Expected exactly one dimension tag.')
	}
	dimension_string := dimension_tags[0].attributes['ref'] or {
		return error('Dimension does not include location.')
	}
	dimension_parts := dimension_string.split(':')
	top_left := Location.from_encoding(dimension_parts[0])!
	bottom_right_code := if dimension_parts.len == 2 {
		dimension_parts[1]
	} else {
		dimension_parts[0]
	}
	mut bottom_right := Location.from_encoding(bottom_right_code)!

	row_tags := doc.get_elements_by_tag('row')

	mut rows := []Row{}

	// Map to store shared formulas by their si index
	mut shared_formulas := map[int]string{}

	for row in row_tags {
		// Get the location of the row.
		row_label := row.attributes['r'] or { return error('Row does not include location.') }
		row_index := row_label.int() - 1

		span_string := row.attributes['spans'] or { '1:1' }

		span := span_string.split(':').map(it.int())
		cell_count := if span.len >= 2 { span[1] - span[0] + 1 } else { 1 }

		mut cells := []Cell{cap: cell_count}

		for child in row.children {
			match child {
				xml.XMLNode {
					// Extract value from <v> element
					matching_tags := child.children.filter(it is xml.XMLNode && it.name == 'v').map(it as xml.XMLNode)
					if matching_tags.len > 1 {
						return error('Expected only one <v> element in cell, found ${matching_tags.len}')
					}
					if matching_tags.len == 0 {
						// Cell with no value (empty cell or styled empty cell) - skip it
						continue
					}
					value_tag := matching_tags[0]

					cell_type := CellType.from_code(child.attributes['t'] or { 'n' })!
					value := if value_tag.children.len == 0 {
						'' // Empty value
					} else if cell_type == .string_type {
						idx := (value_tag.children[0] as string).int()
						if idx >= shared_strings.len {
							return error('Invalid shared string index ${idx}, only ${shared_strings.len} strings available')
						}
						shared_strings[idx]
					} else {
						value_tag.children[0] as string
					}

					// Extract formula from <f> element if present
					formula_tags := child.children.filter(it is xml.XMLNode && it.name == 'f').map(it as xml.XMLNode)
					mut formula := ''
					if formula_tags.len > 0 {
						f_tag := formula_tags[0]
						// Check if this is a shared formula definition or reference
						is_shared := f_tag.attributes['t'] or { '' } == 'shared'
						si_str := f_tag.attributes['si'] or { '' }
						si := if si_str != '' { si_str.int() } else { -1 }

						if f_tag.children.len > 0 {
							// Has formula text - this is either a regular formula or a shared formula definition
							formula = f_tag.children[0] as string
							// If it's a shared formula definition, store it
							if is_shared && si >= 0 {
								shared_formulas[si] = formula
							}
						} else if is_shared && si >= 0 {
							// This is a shared formula reference - look up the formula
							// Note: Ideally we'd adjust cell references, but for now use the base formula
							formula = shared_formulas[si] or { '' }
						}
					}

					location_string := child.attributes['r'] or {
						return error('Cell does not include location reference (r attribute)')
					}

					// Extract style information
					style_id := (child.attributes['s'] or { '0' }).int()

					// Look up currency from style
					mut currency := ?Currency(none)
					mut fill := ?ThemeFill(none)

					if style_id < style_info.cell_xfs.len {
						cell_xf := style_info.cell_xfs[style_id]

						// Check for currency format
						if format_code := style_info.num_fmt_codes[cell_xf.num_fmt_id] {
							currency = detect_currency_from_format(format_code)
						}

						// Check for theme fill (fill_id >= 2 means non-default fill)
						if cell_xf.fill_id >= 2 && cell_xf.fill_id < style_info.fills.len {
							fill_info := style_info.fills[cell_xf.fill_id]
							if theme := fill_info.theme {
								fill = ThemeFill{
									theme: theme
									tint:  fill_info.tint or { 0.0 }
								}
							}
						}
					}

					cells << Cell{
						value:     value
						cell_type: cell_type
						location:  Location.from_encoding(location_string)!
						formula:   formula
						style_id:  style_id
						currency:  currency
						fill:      fill
					}
				}
				else {
					// Non-XML node children (whitespace, etc.) are ignored
					continue
				}
			}
		}

		rows << Row{
			row_index: row_index
			row_label: row_label
			cells:     cells
		}
	}
	return Sheet{
		name:         name
		rows:         rows
		top_left:     top_left
		bottom_right: bottom_right
	}
}
