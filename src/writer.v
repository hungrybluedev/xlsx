module xlsx

import os
import compress.szip
import rand
import strings

// StyleKey represents a unique style combination (numFmtId + fillId)
struct StyleKey {
	num_fmt_id int
	fill_id    int
}

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
	os.mkdir_all(os.join_path(temp_dir, 'xl', 'theme'))!

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

	// Write theme (required for consistent theme colors)
	os.write_file(os.join_path(temp_dir, 'xl', 'theme', 'theme1.xml'), generate_theme_xml())!

	// Collect all unique fills and currencies used in the document
	// Build maps: fill_key -> fillId, format_code -> numFmtId
	// Then build style combinations: (numFmtId, fillId) -> styleId
	mut fills := []ThemeFill{}
	mut fill_map := map[string]int{} // fill_key -> fillId (starts at 2)
	mut num_fmts := []string{}
	mut num_fmt_map := map[string]int{} // format_code -> numFmtId (starts at 164)

	// First pass: collect unique fills and number formats
	for _, sheet in doc.sheets {
		for row in sheet.rows {
			for cell in row.cells {
				// Collect fills
				if fill := cell.fill {
					key := fill.to_key()
					if key !in fill_map {
						fill_map[key] = fills.len + 2 // fillId 0,1 are reserved
						fills << fill
					}
				}
				// Collect number formats (currencies)
				if currency := cell.currency {
					format_code := currency.format_code()
					if format_code !in num_fmt_map {
						num_fmt_map[format_code] = 164 + num_fmts.len
						num_fmts << format_code
					}
				}
			}
		}
	}

	// Build style combinations map: "numFmtId:fillId" -> styleId
	// Pre-defined styles: 0=default, 1=date (numFmtId=16, fillId=0)
	mut style_map := map[string]int{}
	mut styles := []StyleKey{}

	// Add default style (styleId=0)
	styles << StyleKey{
		num_fmt_id: 0
		fill_id:    0
	}
	style_map['0:0'] = 0

	// Add date style (styleId=1, numFmtId=16, fillId=0)
	styles << StyleKey{
		num_fmt_id: 16
		fill_id:    0
	}
	style_map['16:0'] = 1

	// Second pass: build unique style combinations
	for _, sheet in doc.sheets {
		for row in sheet.rows {
			for cell in row.cells {
				// Determine numFmtId
				mut num_fmt_id := 0
				if cell.style_id == 1 {
					num_fmt_id = 16 // Date format
				}
				if currency := cell.currency {
					num_fmt_id = num_fmt_map[currency.format_code()]
				}

				// Determine fillId
				mut fill_id := 0
				if fill := cell.fill {
					fill_id = fill_map[fill.to_key()]
				}

				// Register style combination
				style_key := '${num_fmt_id}:${fill_id}'
				if style_key !in style_map {
					style_map[style_key] = styles.len
					styles << StyleKey{
						num_fmt_id: num_fmt_id
						fill_id:    fill_id
					}
				}
			}
		}
	}

	// Write styles.xml with dynamic fills and Aptos font
	os.write_file(os.join_path(temp_dir, 'xl', 'styles.xml'), generate_styles_v2(num_fmts,
		fills, styles))!

	// Write each sheet
	for sheet_id, sheet in doc.sheets {
		sheet_xml := generate_sheet_xml_v2(sheet, string_index_map, num_fmt_map, fill_map,
			style_map)
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

	sb.write_string('<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>')
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

	// Theme (always included for consistent colors)
	sb.write_string('<Relationship Id="rId${next_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>')
	next_id++

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

// Generate xl/styles.xml with dynamic currency formats
// currency_style_map: maps format_code -> style_id (style_id starts at 2)
fn generate_styles(currency_style_map map[string]int) string {
	mut sb := strings.new_builder(1024)
	sb.write_string('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
	sb.write_string('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')

	// Custom number formats section (only if we have currencies)
	// Custom numFmtIds start at 164
	if currency_style_map.len > 0 {
		sb.write_string('<numFmts count="${currency_style_map.len}">')
		mut num_fmt_id := 164
		for format_code, _ in currency_style_map {
			// Escape special XML characters in format code
			escaped_code := xml_escape(format_code)
			sb.write_string('<numFmt numFmtId="${num_fmt_id}" formatCode="${escaped_code}"/>')
			num_fmt_id++
		}
		sb.write_string('</numFmts>')
	}

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

	// Cell formats (cellXfs)
	// Index 0 = default, Index 1 = date, Index 2+ = currencies
	cell_xf_count := 2 + currency_style_map.len
	sb.write_string('<cellXfs count="${cell_xf_count}">')
	sb.write_string('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>') // s="0" default
	sb.write_string('<xf numFmtId="16" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>') // s="1" date

	// Add currency formats in order of their style_id (which matches iteration order)
	mut num_fmt_id := 164
	for _, _ in currency_style_map {
		sb.write_string('<xf numFmtId="${num_fmt_id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>')
		num_fmt_id++
	}
	sb.write_string('</cellXfs>')

	// Cell styles
	sb.write_string('<cellStyles count="1">')
	sb.write_string('<cellStyle name="Normal" xfId="0" builtinId="0"/>')
	sb.write_string('</cellStyles>')

	sb.write_string('</styleSheet>')
	return sb.str()
}

// Generate xl/worksheets/sheet{N}.xml
// currency_style_map: maps format_code -> style_id for currency cells
fn generate_sheet_xml(sheet Sheet, string_index_map map[string]int, currency_style_map map[string]int) string {
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

			// Determine style_id: currency takes precedence, then style_id field
			mut effective_style_id := cell.style_id
			if currency := cell.currency {
				format_code := currency.format_code()
				if style_id := currency_style_map[format_code] {
					effective_style_id = style_id
				}
			}

			style_attr := if effective_style_id > 0 { ' s="${effective_style_id}"' } else { '' }

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
				// Number cell (with optional style for dates, currencies, etc.)
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

// Generate xl/theme/theme1.xml - Office theme with Aptos fonts
// This defines the actual RGB colors for theme indices used in fills
fn generate_theme_xml() string {
	mut sb := strings.new_builder(4096)
	sb.write_string('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
	sb.write_string('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">')
	sb.write_string('<a:themeElements>')
	// Color scheme - defines theme indices 0-11
	sb.write_string('<a:clrScheme name="Office">')
	sb.write_string('<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>')
	sb.write_string('<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>')
	sb.write_string('<a:dk2><a:srgbClr val="0E2841"/></a:dk2>')
	sb.write_string('<a:lt2><a:srgbClr val="E8E8E8"/></a:lt2>')
	sb.write_string('<a:accent1><a:srgbClr val="156082"/></a:accent1>')
	sb.write_string('<a:accent2><a:srgbClr val="E97132"/></a:accent2>')
	sb.write_string('<a:accent3><a:srgbClr val="196B24"/></a:accent3>')
	sb.write_string('<a:accent4><a:srgbClr val="0F9ED5"/></a:accent4>')
	sb.write_string('<a:accent5><a:srgbClr val="A02B93"/></a:accent5>')
	sb.write_string('<a:accent6><a:srgbClr val="4EA72E"/></a:accent6>')
	sb.write_string('<a:hlink><a:srgbClr val="467886"/></a:hlink>')
	sb.write_string('<a:folHlink><a:srgbClr val="96607D"/></a:folHlink>')
	sb.write_string('</a:clrScheme>')
	// Font scheme - Aptos fonts
	sb.write_string('<a:fontScheme name="Office">')
	sb.write_string('<a:majorFont><a:latin typeface="Aptos Display" panose="02110004020202020204"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>')
	sb.write_string('<a:minorFont><a:latin typeface="Aptos Narrow" panose="02110004020202020204"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>')
	sb.write_string('</a:fontScheme>')
	// Format scheme
	sb.write_string('<a:fmtScheme name="Office">')
	sb.write_string('<a:fillStyleLst>')
	sb.write_string('<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>')
	sb.write_string('<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>')
	sb.write_string('<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>')
	sb.write_string('</a:fillStyleLst>')
	sb.write_string('<a:lnStyleLst>')
	sb.write_string('<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>')
	sb.write_string('<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>')
	sb.write_string('<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>')
	sb.write_string('</a:lnStyleLst>')
	sb.write_string('<a:effectStyleLst>')
	sb.write_string('<a:effectStyle><a:effectLst/></a:effectStyle>')
	sb.write_string('<a:effectStyle><a:effectLst/></a:effectStyle>')
	sb.write_string('<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>')
	sb.write_string('</a:effectStyleLst>')
	sb.write_string('<a:bgFillStyleLst>')
	sb.write_string('<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>')
	sb.write_string('<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>')
	sb.write_string('<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>')
	sb.write_string('</a:bgFillStyleLst>')
	sb.write_string('</a:fmtScheme>')
	sb.write_string('</a:themeElements>')
	sb.write_string('<a:objectDefaults/>')
	sb.write_string('<a:extraClrSchemeLst/>')
	sb.write_string('</a:theme>')
	return sb.str()
}

// Generate xl/styles.xml with dynamic fills, number formats, and Aptos font
fn generate_styles_v2(num_fmts []string, fills []ThemeFill, styles []StyleKey) string {
	mut sb := strings.new_builder(2048)
	sb.write_string('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
	sb.write_string('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')

	// Custom number formats section (only if we have any)
	// Custom numFmtIds start at 164
	if num_fmts.len > 0 {
		sb.write_string('<numFmts count="${num_fmts.len}">')
		for i, format_code in num_fmts {
			num_fmt_id := 164 + i
			escaped_code := xml_escape(format_code)
			sb.write_string('<numFmt numFmtId="${num_fmt_id}" formatCode="${escaped_code}"/>')
		}
		sb.write_string('</numFmts>')
	}

	// Fonts - Aptos Narrow 12pt (Microsoft's new default font)
	sb.write_string('<fonts count="1">')
	sb.write_string('<font>')
	sb.write_string('<sz val="12"/>')
	sb.write_string('<color theme="1"/>')
	sb.write_string('<name val="Aptos Narrow"/>')
	sb.write_string('<family val="2"/>')
	sb.write_string('<scheme val="minor"/>')
	sb.write_string('</font>')
	sb.write_string('</fonts>')

	// Fills - two required (none and gray125) plus custom fills
	fill_count := 2 + fills.len
	sb.write_string('<fills count="${fill_count}">')
	sb.write_string('<fill><patternFill patternType="none"/></fill>')
	sb.write_string('<fill><patternFill patternType="gray125"/></fill>')
	for fill in fills {
		sb.write_string('<fill><patternFill patternType="solid">')
		sb.write_string('<fgColor theme="${fill.theme}" tint="${fill.tint}"/>')
		sb.write_string('<bgColor indexed="64"/>')
		sb.write_string('</patternFill></fill>')
	}
	sb.write_string('</fills>')

	// Borders - at least one required
	sb.write_string('<borders count="1">')
	sb.write_string('<border><left/><right/><top/><bottom/><diagonal/></border>')
	sb.write_string('</borders>')

	// Cell style formats
	sb.write_string('<cellStyleXfs count="1">')
	sb.write_string('<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>')
	sb.write_string('</cellStyleXfs>')

	// Cell formats (cellXfs) - all unique style combinations
	sb.write_string('<cellXfs count="${styles.len}">')
	for style in styles {
		mut attrs := 'numFmtId="${style.num_fmt_id}" fontId="0" fillId="${style.fill_id}" borderId="0" xfId="0"'
		if style.num_fmt_id != 0 {
			attrs += ' applyNumberFormat="1"'
		}
		if style.fill_id > 1 {
			attrs += ' applyFill="1"'
		}
		sb.write_string('<xf ${attrs}/>')
	}
	sb.write_string('</cellXfs>')

	// Cell styles
	sb.write_string('<cellStyles count="1">')
	sb.write_string('<cellStyle name="Normal" xfId="0" builtinId="0"/>')
	sb.write_string('</cellStyles>')

	sb.write_string('</styleSheet>')
	return sb.str()
}

// Generate xl/worksheets/sheet{N}.xml with fill support
fn generate_sheet_xml_v2(sheet Sheet, string_index_map map[string]int, num_fmt_map map[string]int, fill_map map[string]int, style_map map[string]int) string {
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

			// Determine numFmtId
			mut num_fmt_id := 0
			if cell.style_id == 1 {
				num_fmt_id = 16 // Date format
			}
			if currency := cell.currency {
				num_fmt_id = num_fmt_map[currency.format_code()]
			}

			// Determine fillId
			mut fill_id := 0
			if fill := cell.fill {
				fill_id = fill_map[fill.to_key()]
			}

			// Look up effective style_id from style_map
			style_key := '${num_fmt_id}:${fill_id}'
			effective_style_id := style_map[style_key]

			style_attr := if effective_style_id > 0 { ' s="${effective_style_id}"' } else { '' }

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
				// Number cell (with optional style for dates, currencies, etc.)
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
