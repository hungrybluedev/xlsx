module main

import os
import time
import xlsx

// Test the example from README.md to ensure documentation stays accurate
fn test_readme_writer_example() ! {
	// Create a new document
	mut doc := xlsx.Document.new()

	// Add a sheet and get a mutable reference
	sheet_id := doc.add_sheet('Sheet1')
	mut sheet := doc.get_sheet_mut(sheet_id) or { return error('Failed to get sheet') }

	// Set string cells
	sheet.set_cell(xlsx.Location.from_encoding('A1')!, 'Name')
	sheet.set_cell(xlsx.Location.from_encoding('B1')!, 'Score')

	// Set numeric cells
	sheet.set_number(xlsx.Location.from_encoding('B2')!, 95)
	sheet.set_number_f64(xlsx.Location.from_encoding('B3')!, 87.5)

	// Set dates using time.Time
	jan_1_2026 := time.Time{
		year:  2026
		month: 1
		day:   1
	}
	sheet.set_date(xlsx.Location.from_encoding('C2')!, jan_1_2026)

	// Set formulas
	sheet.set_formula(xlsx.Location.from_encoding('B4')!, 'SUM(B2:B3)')

	// Set currency values (USD format)
	sheet.set_currency(xlsx.Location.from_encoding('D2')!, 1234.56, .usd)

	// Save to file
	temp_path := os.join_path(os.temp_dir(), 'test_readme_example.xlsx')
	defer {
		os.rm(temp_path) or {}
	}

	doc.to_file(temp_path)!

	// Verify the file was created and is a valid ZIP (XLSX)
	assert os.exists(temp_path), 'Output file should exist'

	content := os.read_bytes(temp_path) or { return error('Failed to read output file') }
	assert content.len > 0, 'Output file should not be empty'
	// ZIP files start with PK (0x50 0x4B)
	assert content[0] == 0x50, 'File should start with ZIP signature'
	assert content[1] == 0x4B, 'File should start with ZIP signature'

	// Roundtrip: read the file back and verify content
	read_doc := xlsx.Document.from_file(temp_path)!
	assert read_doc.sheets.len == 1, 'Should have 1 sheet'

	read_sheet := read_doc.sheets[1]
	assert read_sheet.name == 'Sheet1', 'Sheet name should be Sheet1'

	data := read_sheet.get_all_data()!
	assert data.raw_data.len >= 4, 'Should have at least 4 rows'

	// Verify header cells
	assert data.raw_data[0][0] == 'Name', 'A1 should be Name'
	assert data.raw_data[0][1] == 'Score', 'B1 should be Score'
}

// Test the new CellBuilder API
fn test_cell_builder_api() ! {
	mut doc := xlsx.Document.new()
	sheet_id := doc.add_sheet('BuilderTest')
	mut sheet := doc.get_sheet_mut(sheet_id) or { return error('Failed to get sheet') }

	// Test simple string cell with build_cell
	sheet.build_cell(xlsx.Location.from_encoding('A1')!, text: 'Hello')

	// Test number with fill
	fill := xlsx.ThemeFill{
		theme: 4
		tint:  0.6
	}
	sheet.build_cell(xlsx.Location.from_encoding('B1')!, number: 42.5, fill: fill)

	// Test formula with currency
	sheet.build_cell(xlsx.Location.from_encoding('C1')!, formula: 'A1+B1', currency: .usd)

	// Verify cells were created
	assert sheet.rows.len == 1, 'Should have 1 row'
	assert sheet.rows[0].cells.len == 3, 'Should have 3 cells'

	// Save and verify
	temp_path := os.join_path(os.temp_dir(), 'test_cell_builder.xlsx')
	defer {
		os.rm(temp_path) or {}
	}

	doc.to_file(temp_path)!
	assert os.exists(temp_path), 'Output file should exist'
}
