module xlsx

import os
import compress.szip
import rand
import strings

// Creates a new empty Document for writing
pub fn Document.new() Document {
	return Document{
		shared_strings: []string{}
		sheets:         map[int]Sheet{}
	}
}

// Adds a new sheet to the document and returns its ID
pub fn (mut doc Document) add_sheet(name string) int {
	sheet_id := doc.sheets.len + 1
	doc.sheets[sheet_id] = Sheet{
		name: name
		rows: []Row{}
	}
	return sheet_id
}

// Gets a mutable reference to a sheet by its ID
pub fn (mut doc Document) get_sheet_mut(sheet_id int) ?&Sheet {
	if sheet_id in doc.sheets {
		return unsafe { &doc.sheets[sheet_id] }
	}
	return none
}

// Writes the document to an XLSX file
pub fn (doc Document) to_file(path string) ! {
	// Create a temporary directory for the XLSX contents
	temp_dir := os.join_path(os.temp_dir(), 'xlsx_writer_${os.getpid()}_${rand_string(8)}')
	os.mkdir_all(temp_dir) or { return error('Failed to create temp directory: ${err}') }
	defer {
		os.rmdir_all(temp_dir) or {}
	}

	// Build shared strings table from all cells
	mut shared_strings := []string{}
	mut string_index_map := map[string]int{}
	for _, sheet in doc.sheets {
		for row in sheet.rows {
			for cell in row.cells {
				if cell.cell_type == .string_type && cell.value !in string_index_map {
					string_index_map[cell.value] = shared_strings.len
					shared_strings << cell.value
				}
			}
		}
	}

	// Create directory structure
	os.mkdir_all(os.join_path(temp_dir, '_rels'))!
	os.mkdir_all(os.join_path(temp_dir, 'xl', '_rels'))!
	os.mkdir_all(os.join_path(temp_dir, 'xl', 'worksheets'))!

	// Generate and write XML files
	os.write_file(os.join_path(temp_dir, '[Content_Types].xml'), generate_content_types(doc,
		shared_strings.len > 0))!
	os.write_file(os.join_path(temp_dir, '_rels', '.rels'), generate_root_rels())!
	os.write_file(os.join_path(temp_dir, 'xl', 'workbook.xml'), generate_workbook(doc))!
	os.write_file(os.join_path(temp_dir, 'xl', '_rels', 'workbook.xml.rels'), generate_workbook_rels(doc,
		shared_strings.len > 0))!

	// Write shared strings if any exist
	if shared_strings.len > 0 {
		os.write_file(os.join_path(temp_dir, 'xl', 'sharedStrings.xml'), generate_shared_strings(shared_strings))!
	}

	// Write minimal styles.xml (required for proper file structure)
	os.write_file(os.join_path(temp_dir, 'xl', 'styles.xml'), generate_styles())!

	// Write each sheet
	for sheet_id, sheet in doc.sheets {
		sheet_xml := generate_sheet_xml(sheet, string_index_map)
		os.write_file(os.join_path(temp_dir, 'xl', 'worksheets', 'sheet${sheet_id}.xml'),
			sheet_xml)!
	}

	// Create ZIP file
	// Remove existing file if present
	if os.exists(path) {
		os.rm(path)!
	}

	// Create the ZIP archive using system zip command for better compatibility
	// (some XLSX readers are sensitive to ZIP implementation details)
	abs_path := os.real_path(path)
	result := os.execute('cd "${temp_dir}" && zip -r "${abs_path}" .')
	if result.exit_code != 0 {
		// Fallback to szip if system zip is not available
		szip.zip_folder(temp_dir, path, szip.ZipFolderOptions{})!
	}
}

// Generate a random string for temp directory names
fn rand_string(len int) string {
	mut result := strings.new_builder(len)
	chars := 'abcdefghijklmnopqrstuvwxyz0123456789'
	for _ in 0 .. len {
		result.write_u8(chars[rand.intn(chars.len) or { 0 }])
	}
	return result.str()
}

// Generate [Content_Types].xml
fn generate_content_types(doc Document, has_shared_strings bool) string {
	mut sb := strings.new_builder(512)
	sb.write_string('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
	sb.write_string('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
	sb.write_string('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
	sb.write_string('<Default Extension="xml" ContentType="application/xml"/>')
	sb.write_string('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')

	for sheet_id, _ in doc.sheets {
		sb.write_string('<Override PartName="/xl/worksheets/sheet${sheet_id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
	}

	sb.write_string('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')

	if has_shared_strings {
		sb.write_string('<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>')
	}

	sb.write_string('</Types>')
	return sb.str()
}

// Generate _rels/.rels
fn generate_root_rels() string {
	mut sb := strings.new_builder(256)
	sb.write_string('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
	sb.write_string('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
	sb.write_string('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>')
	sb.write_string('</Relationships>')
	return sb.str()
}

// Generate xl/workbook.xml
fn generate_workbook(doc Document) string {
	mut sb := strings.new_builder(512)
	sb.write_string('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
	sb.write_string('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')

	// Book views (required for proper display)
	sb.write_string('<bookViews>')
	sb.write_string('<workbookView xWindow="0" yWindow="0" windowWidth="16384" windowHeight="8192"/>')
	sb.write_string('</bookViews>')

	sb.write_string('<sheets>')
	for sheet_id, sheet in doc.sheets {
		sb.write_string('<sheet name="${xml_escape(sheet.name)}" sheetId="${sheet_id}" r:id="rId${sheet_id}"/>')
	}
	sb.write_string('</sheets>')

	sb.write_string('</workbook>')
	return sb.str()
}

// Generate xl/_rels/workbook.xml.rels
fn generate_workbook_rels(doc Document, has_shared_strings bool) string {
	mut sb := strings.new_builder(512)
	sb.write_string('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
	sb.write_string('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')

	mut next_id := 1

	// Worksheets
	for sheet_id, _ in doc.sheets {
		sb.write_string('<Relationship Id="rId${next_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${sheet_id}.xml"/>')
		next_id++
	}

	// Styles (always included)
	sb.write_string('<Relationship Id="rId${next_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
	next_id++

	// Shared strings (if any)
	if has_shared_strings {
		sb.write_string('<Relationship Id="rId${next_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>')
	}

	sb.write_string('</Relationships>')
	return sb.str()
}

// Generate xl/sharedStrings.xml
fn generate_shared_strings(strings_list []string) string {
	mut sb := strings.new_builder(1024)
	sb.write_string('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
	sb.write_string('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings_list.len}" uniqueCount="${strings_list.len}">')

	for s in strings_list {
		sb.write_string('<si><t>${xml_escape(s)}</t></si>')
	}

	sb.write_string('</sst>')
	return sb.str()
}

// Generate xl/styles.xml (minimal required structure)
fn generate_styles() string {
	mut sb := strings.new_builder(1024)
	sb.write_string('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
	sb.write_string('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')

	// Fonts - at least one required
	sb.write_string('<fonts count="1">')
	sb.write_string('<font><sz val="11"/><name val="Calibri"/></font>')
	sb.write_string('</fonts>')

	// Fills - at least two required (none and gray125)
	sb.write_string('<fills count="2">')
	sb.write_string('<fill><patternFill patternType="none"/></fill>')
	sb.write_string('<fill><patternFill patternType="gray125"/></fill>')
	sb.write_string('</fills>')

	// Borders - at least one required
	sb.write_string('<borders count="1">')
	sb.write_string('<border><left/><right/><top/><bottom/><diagonal/></border>')
	sb.write_string('</borders>')

	// Cell style formats
	sb.write_string('<cellStyleXfs count="1">')
	sb.write_string('<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>')
	sb.write_string('</cellStyleXfs>')

	// Cell formats - index 0 is default, index 1 is date format, index 2 is GBP currency
	// numFmtId 16 is built-in "d-mmm" format (e.g., "1-Jan")
	// numFmtId 8 is built-in GBP currency format (e.g., "£17.80")
	sb.write_string('<cellXfs count="3">')
	sb.write_string('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>') // s="0" default
	sb.write_string('<xf numFmtId="16" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>') // s="1" date
	sb.write_string('<xf numFmtId="8" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>') // s="2" GBP currency
	sb.write_string('</cellXfs>')

	// Cell styles
	sb.write_string('<cellStyles count="1">')
	sb.write_string('<cellStyle name="Normal" xfId="0" builtinId="0"/>')
	sb.write_string('</cellStyles>')

	sb.write_string('</styleSheet>')
	return sb.str()
}

// Generate xl/worksheets/sheet{N}.xml
fn generate_sheet_xml(sheet Sheet, string_index_map map[string]int) string {
	mut sb := strings.new_builder(2048)
	sb.write_string('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
	sb.write_string('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')

	// Calculate dimension and global spans
	mut dim_ref := 'A1'
	mut global_min_col := 0
	mut global_max_col := 0

	if sheet.rows.len > 0 {
		mut min_row := 999999
		mut max_row := 0
		global_min_col = 999999
		global_max_col = 0

		for row in sheet.rows {
			if row.row_index < min_row {
				min_row = row.row_index
			}
			if row.row_index > max_row {
				max_row = row.row_index
			}
			for cell in row.cells {
				if cell.location.col < global_min_col {
					global_min_col = cell.location.col
				}
				if cell.location.col > global_max_col {
					global_max_col = cell.location.col
				}
			}
		}

		top_left := Location.from_cartesian(min_row, global_min_col) or { Location{} }
		bottom_right := Location.from_cartesian(max_row, global_max_col) or { Location{} }
		dim_ref = '${top_left.col_label}${top_left.row_label}:${bottom_right.col_label}${bottom_right.row_label}'
	}

	// Use consistent spans for all rows (based on sheet dimension)
	global_spans := '${global_min_col + 1}:${global_max_col + 1}'

	sb.write_string('<dimension ref="${dim_ref}"/>')

	// Sheet views
	sb.write_string('<sheetViews>')
	sb.write_string('<sheetView tabSelected="1" workbookViewId="0"/>')
	sb.write_string('</sheetViews>')

	// Sheet format
	sb.write_string('<sheetFormatPr defaultRowHeight="15"/>')

	sb.write_string('<sheetData>')

	// Sort rows by row_index
	mut sorted_rows := sheet.rows.clone()
	sorted_rows.sort(a.row_index < b.row_index)

	for row in sorted_rows {
		// Skip empty rows
		if row.cells.len == 0 {
			continue
		}

		row_num := row.row_index + 1 // Excel uses 1-based rows

		sb.write_string('<row r="${row_num}" spans="${global_spans}">')

		// Sort cells by column
		mut sorted_cells := row.cells.clone()
		sorted_cells.sort(a.location.col < b.location.col)

		for cell in sorted_cells {
			cell_ref := '${cell.location.col_label}${row_num}'
			style_attr := if cell.style_id > 0 { ' s="${cell.style_id}"' } else { '' }

			if cell.formula.len > 0 {
				// Formula cell - include placeholder value (Excel will recalculate)
				sb.write_string('<c r="${cell_ref}"${style_attr}><f>${xml_escape(cell.formula)}</f>')
				if cell.value.len > 0 {
					sb.write_string('<v>${cell.value}</v>')
				} else {
					sb.write_string('<v>0</v>')
				}
				sb.write_string('</c>')
			} else if cell.cell_type == .string_type {
				// String cell - reference shared strings
				idx := string_index_map[cell.value]
				sb.write_string('<c r="${cell_ref}"${style_attr} t="s"><v>${idx}</v></c>')
			} else {
				// Number cell (with optional style for dates, etc.)
				sb.write_string('<c r="${cell_ref}"${style_attr}><v>${cell.value}</v></c>')
			}
		}

		sb.write_string('</row>')
	}

	sb.write_string('</sheetData>')

	// Page margins (commonly expected)
	sb.write_string('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')

	sb.write_string('</worksheet>')
	return sb.str()
}

// Escape XML special characters
fn xml_escape(s string) string {
	return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"',
		'&quot;').replace("'", '&apos;')
}
