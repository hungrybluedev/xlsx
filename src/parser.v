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

	if os.exists(shared_strings_path) {
		strings_doc := xml.XMLDocument.from_file(shared_strings_path) or {
			return error('Failed to parse shared strings file of excel file: ${path}')
		}

		all_defined_strings := strings_doc.get_elements_by_tag('si')
		for definition in all_defined_strings {
			t_element := definition.children[0]
			if t_element !is xml.XMLNode || (t_element as xml.XMLNode).name != 't' {
				return error('Invalid shared string definition: ${definition}')
			}

			content := (t_element as xml.XMLNode).children[0]
			if content !is string {
				return error('Invalid shared string definition: ${definition}')
			}
			shared_strings << (content as string)
		}
	}
	return shared_strings
}

fn load_worksheets_metadata(path string, worksheets_file_path string) !map[int]string {
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

pub fn Document.from_file(path string) !Document {
	// Fail if the file does not exist.
	if !os.exists(path) {
		return error('File does not exist: ${path}')
	}
	// First, we extract the ZIP file into a temporary directory.
	location := create_temporary_directory()

	szip.extract_zip_to_dir(path, location) or {
		return error('Failed to extract information from file: ${path}')
	}

	// Then we list the files in the "xl" directory.
	xl_path := os.join_path(location, 'xl')

	// Load the strings from the shared strings file, if it exists.
	shared_strings_path := os.join_path(xl_path, 'sharedStrings.xml')
	shared_strings := load_shared_strings(path, shared_strings_path)!

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

		sheet := Sheet.from_doc(sheet_name, sheet_doc, shared_strings) or {
			return error('Failed to parse sheet file: ${sheet_path}')
		}

		sheet_map[sheet_id] = sheet
	}

	return Document{
		shared_strings: shared_strings
		sheets: sheet_map
	}
}

fn Sheet.from_doc(name string, doc xml.XMLDocument, shared_strings []string) !Sheet {
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

	row_loop: for row in row_tags {
		// Get the location of the row.
		row_label := row.attributes['r'] or { return error('Row does not include location.') }
		row_index := row_label.int() - 1

		span_string := row.attributes['spans'] or { return error('Row does not include span.') }

		span := span_string.split(':').map(it.int())
		cell_count := span[1] - span[0] + 1

		mut cells := []Cell{cap: cell_count}

		for child in row.children {
			match child {
				xml.XMLNode {
					// First, we check if the cell is empty
					if child.children.len == 0 {
						bottom_right = Location.from_cartesian(row_index - 1, bottom_right.col)!
						break row_loop
					}
					matching_tags := child.children.filter(it is xml.XMLNode && it.name == 'v').map(it as xml.XMLNode)
					if matching_tags.len > 1 {
						return error('Expected only one value: ${child}')
					}
					value_tag := matching_tags[0]

					cell_type := CellType.from_code(child.attributes['t'] or { 'n' })!
					value := if cell_type == .string_type {
						shared_strings[(value_tag.children[0] as string).int()]
					} else {
						value_tag.children[0] as string
					}

					location_string := child.attributes['r'] or {
						return error('Cell does not include location.')
					}

					cells << Cell{
						value: value
						cell_type: cell_type
						location: Location.from_encoding(location_string)!
					}
				}
				else {
					return error('Invalid cell of row: ${child}')
				}
			}
		}

		rows << Row{
			row_index: row_index
			row_label: row_label
			cells: cells
		}
	}

	return Sheet{
		name: name
		rows: rows
		top_left: top_left
		bottom_right: bottom_right
	}
}
