module xlsx

import os

// Unit tests for XLSX writer functionality

fn test_document_new_creates_empty_document() {
	doc := Document.new()

	assert doc.shared_strings.len == 0, 'shared_strings should be empty'
	assert doc.sheets.len == 0, 'sheets should be empty'
}

fn test_document_add_sheet_creates_sheet() {
	mut doc := Document.new()

	sheet_id := doc.add_sheet('TestSheet')

	assert sheet_id == 1, 'first sheet should have ID 1'
	assert doc.sheets.len == 1, 'should have one sheet'
	assert doc.sheets[1].name == 'TestSheet', 'sheet name should match'
}

fn test_document_add_multiple_sheets() {
	mut doc := Document.new()

	id1 := doc.add_sheet('Sheet1')
	id2 := doc.add_sheet('Sheet2')

	assert id1 == 1
	assert id2 == 2
	assert doc.sheets.len == 2
	assert doc.sheets[1].name == 'Sheet1'
	assert doc.sheets[2].name == 'Sheet2'
}

fn test_sheet_set_cell_stores_string() ! {
	mut doc := Document.new()
	sheet_id := doc.add_sheet('Test')
	mut sheet := doc.get_sheet_mut(sheet_id)?

	loc := Location.from_encoding('A1')!
	sheet.set_cell(loc, 'Hello')

	assert sheet.rows.len == 1, 'should have one row'
	assert sheet.rows[0].cells.len == 1, 'should have one cell'
	assert sheet.rows[0].cells[0].value == 'Hello', 'cell value should match'
	assert sheet.rows[0].cells[0].cell_type == .string_type, 'should be string type'
}

fn test_sheet_set_number_stores_int() ! {
	mut doc := Document.new()
	sheet_id := doc.add_sheet('Test')
	mut sheet := doc.get_sheet_mut(sheet_id)?

	loc := Location.from_encoding('B2')!
	sheet.set_number(loc, 42)

	assert sheet.rows.len == 1
	assert sheet.rows[0].cells.len == 1
	assert sheet.rows[0].cells[0].value == '42'
	assert sheet.rows[0].cells[0].cell_type == .number_type
}

fn test_sheet_set_number_f64_stores_float() ! {
	mut doc := Document.new()
	sheet_id := doc.add_sheet('Test')
	mut sheet := doc.get_sheet_mut(sheet_id)?

	loc := Location.from_encoding('C3')!
	sheet.set_number_f64(loc, 3.14)

	assert sheet.rows.len == 1
	assert sheet.rows[0].cells.len == 1
	assert sheet.rows[0].cells[0].value == '3.14'
	assert sheet.rows[0].cells[0].cell_type == .number_type
}

fn test_sheet_set_formula_stores_formula() ! {
	mut doc := Document.new()
	sheet_id := doc.add_sheet('Test')
	mut sheet := doc.get_sheet_mut(sheet_id)?

	loc := Location.from_encoding('D4')!
	sheet.set_formula(loc, 'A1+B1')

	assert sheet.rows.len == 1
	assert sheet.rows[0].cells.len == 1
	assert sheet.rows[0].cells[0].formula == 'A1+B1'
	assert sheet.rows[0].cells[0].cell_type == .number_type
}

fn test_sheet_set_date_stores_date_with_style() ! {
	mut doc := Document.new()
	sheet_id := doc.add_sheet('Test')
	mut sheet := doc.get_sheet_mut(sheet_id)?

	loc := Location.from_encoding('E5')!
	// Excel date 46023 = 2026-01-01
	sheet.set_date(loc, 46023)

	assert sheet.rows.len == 1
	assert sheet.rows[0].cells.len == 1
	assert sheet.rows[0].cells[0].value == '46023', 'date should be stored as number'
	assert sheet.rows[0].cells[0].cell_type == .number_type, 'date should be number type'
	assert sheet.rows[0].cells[0].style_id == 1, 'date should have style_id=1 for date formatting'
}

fn test_sheet_multiple_cells_same_row() ! {
	mut doc := Document.new()
	sheet_id := doc.add_sheet('Test')
	mut sheet := doc.get_sheet_mut(sheet_id)?

	sheet.set_cell(Location.from_encoding('A1')!, 'Name')
	sheet.set_cell(Location.from_encoding('B1')!, 'Value')
	sheet.set_number(Location.from_encoding('C1')!, 100)

	assert sheet.rows.len == 1, 'should still be one row'
	assert sheet.rows[0].cells.len == 3, 'should have three cells'
}

fn test_sheet_multiple_rows() ! {
	mut doc := Document.new()
	sheet_id := doc.add_sheet('Test')
	mut sheet := doc.get_sheet_mut(sheet_id)?

	sheet.set_cell(Location.from_encoding('A1')!, 'Row 1')
	sheet.set_cell(Location.from_encoding('A2')!, 'Row 2')
	sheet.set_cell(Location.from_encoding('A3')!, 'Row 3')

	assert sheet.rows.len == 3, 'should have three rows'
}

fn test_document_to_file_creates_xlsx() ! {
	mut doc := Document.new()
	sheet_id := doc.add_sheet('Sheet1')
	mut sheet := doc.get_sheet_mut(sheet_id)?

	sheet.set_cell(Location.from_encoding('A1')!, 'Hello')
	sheet.set_cell(Location.from_encoding('B1')!, 'World')
	sheet.set_number(Location.from_encoding('A2')!, 42)

	// Write to temp file
	temp_path := os.join_path(os.temp_dir(), 'test_writer_output.xlsx')
	defer {
		os.rm(temp_path) or {}
	}

	doc.to_file(temp_path)!

	// Verify file exists
	assert os.exists(temp_path), 'output file should exist'

	// Verify it's a valid ZIP (XLSX files start with PK)
	content := os.read_bytes(temp_path) or { return error('failed to read file') }
	assert content.len > 4, 'file should have content'
	assert content[0] == 0x50, 'should start with P (ZIP signature)'
	assert content[1] == 0x4B, 'should have K (ZIP signature)'
}

fn test_document_to_file_roundtrip() ! {
	mut doc := Document.new()
	sheet_id := doc.add_sheet('TestSheet')
	mut sheet := doc.get_sheet_mut(sheet_id)?

	sheet.set_cell(Location.from_encoding('A1')!, 'Name')
	sheet.set_cell(Location.from_encoding('B1')!, 'Value')
	sheet.set_cell(Location.from_encoding('A2')!, 'Test')
	sheet.set_number(Location.from_encoding('B2')!, 123)

	// Write to temp file
	temp_path := os.join_path(os.temp_dir(), 'test_roundtrip.xlsx')
	defer {
		os.rm(temp_path) or {}
	}

	doc.to_file(temp_path)!

	// Read it back
	read_doc := Document.from_file(temp_path)!

	assert read_doc.sheets.len == 1, 'should have one sheet'
	read_sheet := read_doc.sheets[1]
	assert read_sheet.name == 'TestSheet', 'sheet name should match'

	// Verify cells
	cell_a1 := read_sheet.get_cell(Location.from_encoding('A1')!)?
	assert cell_a1.value == 'Name', 'A1 should be Name'

	cell_b1 := read_sheet.get_cell(Location.from_encoding('B1')!)?
	assert cell_b1.value == 'Value', 'B1 should be Value'

	cell_a2 := read_sheet.get_cell(Location.from_encoding('A2')!)?
	assert cell_a2.value == 'Test', 'A2 should be Test'

	cell_b2 := read_sheet.get_cell(Location.from_encoding('B2')!)?
	assert cell_b2.value == '123', 'B2 should be 123'
}

// Currency enum tests
fn test_currency_format_code() {
	assert Currency.usd.format_code() == '[$$-409]#,##0.00'
	assert Currency.gbp.format_code() == '[$£-809]#,##0.00'
	assert Currency.eur.format_code() == '[$€-407]#,##0.00'
	assert Currency.jpy.format_code() == '[$¥-411]#,##0' // No decimals for Yen
	assert Currency.cny.format_code() == '[$¥-804]#,##0.00'
	assert Currency.inr.format_code() == '[$₹-4009]#,##0.00'
}

fn test_currency_symbol() {
	assert Currency.usd.symbol() == '$'
	assert Currency.gbp.symbol() == '£'
	assert Currency.eur.symbol() == '€'
	assert Currency.jpy.symbol() == '¥'
}

fn test_sheet_set_currency_with_enum() ! {
	mut doc := Document.new()
	sheet_id := doc.add_sheet('Test')
	mut sheet := doc.get_sheet_mut(sheet_id)?

	loc := Location.from_encoding('C4')!
	sheet.set_currency(loc, 17.80, .gbp)

	assert sheet.rows.len == 1
	assert sheet.rows[0].cells.len == 1
	assert sheet.rows[0].cells[0].value == '17.8'
	assert sheet.rows[0].cells[0].cell_type == .number_type
	assert sheet.rows[0].cells[0].currency? == Currency.gbp, 'should have gbp currency set'
}

fn test_sheet_set_formula_currency() ! {
	mut doc := Document.new()
	sheet_id := doc.add_sheet('Test')
	mut sheet := doc.get_sheet_mut(sheet_id)?

	loc := Location.from_encoding('E4')!
	sheet.set_formula_currency(loc, 'C4*D4', .usd)

	assert sheet.rows.len == 1
	assert sheet.rows[0].cells.len == 1
	assert sheet.rows[0].cells[0].formula == 'C4*D4'
	assert sheet.rows[0].cells[0].cell_type == .number_type
	assert sheet.rows[0].cells[0].currency? == Currency.usd, 'formula should have usd currency'
}

fn test_multiple_currencies_in_document() ! {
	mut doc := Document.new()
	sheet_id := doc.add_sheet('Test')
	mut sheet := doc.get_sheet_mut(sheet_id)?

	sheet.set_currency(Location.from_encoding('A1')!, 100.00, .usd)
	sheet.set_currency(Location.from_encoding('B1')!, 85.00, .gbp)
	sheet.set_currency(Location.from_encoding('C1')!, 92.00, .eur)

	assert sheet.rows.len == 1
	assert sheet.rows[0].cells.len == 3
	assert sheet.rows[0].cells[0].currency? == Currency.usd
	assert sheet.rows[0].cells[1].currency? == Currency.gbp
	assert sheet.rows[0].cells[2].currency? == Currency.eur
}

fn test_sheet_set_formula_with_style_stores_formula_and_style() ! {
	mut doc := Document.new()
	sheet_id := doc.add_sheet('Test')
	mut sheet := doc.get_sheet_mut(sheet_id)?

	loc := Location.from_encoding('E4')!
	sheet.set_formula_with_style(loc, 'C4*D4', 2)

	assert sheet.rows.len == 1
	assert sheet.rows[0].cells.len == 1
	assert sheet.rows[0].cells[0].formula == 'C4*D4'
	assert sheet.rows[0].cells[0].cell_type == .number_type
	assert sheet.rows[0].cells[0].style_id == 2, 'formula should have style_id=2'
}

fn test_theme_fill_creation() {
	fill := ThemeFill{
		theme: 3
		tint:  0.75
	}
	assert fill.theme == 3
	assert fill.tint == 0.75
}

fn test_theme_fill_to_key() {
	fill := ThemeFill{
		theme: 3
		tint:  0.75
	}
	key := fill.to_key()
	assert key == 'theme:3:0.75', 'key should encode theme and tint'
}

fn test_sheet_set_number_with_fill() ! {
	mut doc := Document.new()
	sheet_id := doc.add_sheet('Test')
	mut sheet := doc.get_sheet_mut(sheet_id)?

	loc := Location.from_encoding('D4')!
	fill := ThemeFill{
		theme: 3
		tint:  0.75
	}
	sheet.set_number_with_fill(loc, 39, fill)

	assert sheet.rows.len == 1
	assert sheet.rows[0].cells.len == 1
	assert sheet.rows[0].cells[0].value == '39'
	assert sheet.rows[0].cells[0].cell_type == .number_type
	cell_fill := sheet.rows[0].cells[0].fill or { return error('fill should be set') }
	assert cell_fill.theme == 3, 'fill theme should be 3'
	assert cell_fill.tint == 0.75, 'fill tint should be 0.75'
}

fn test_sheet_set_date_with_fill() ! {
	mut doc := Document.new()
	sheet_id := doc.add_sheet('Test')
	mut sheet := doc.get_sheet_mut(sheet_id)?

	loc := Location.from_encoding('D3')!
	fill := ThemeFill{
		theme: 3
		tint:  0.75
	}
	sheet.set_date_with_fill(loc, 46023, fill)

	assert sheet.rows.len == 1
	assert sheet.rows[0].cells.len == 1
	assert sheet.rows[0].cells[0].value == '46023'
	assert sheet.rows[0].cells[0].style_id == 1, 'date should have style_id=1'
	cell_fill := sheet.rows[0].cells[0].fill or { return error('fill should be set') }
	assert cell_fill.theme == 3
}

fn test_sheet_set_formula_with_fill() ! {
	mut doc := Document.new()
	sheet_id := doc.add_sheet('Test')
	mut sheet := doc.get_sheet_mut(sheet_id)?

	loc := Location.from_encoding('I4')!
	fill := ThemeFill{
		theme: 5
		tint:  0.6
	}
	sheet.set_formula_with_fill(loc, 'IF(D4>40,D4-40,0)', fill)

	assert sheet.rows.len == 1
	assert sheet.rows[0].cells.len == 1
	assert sheet.rows[0].cells[0].formula == 'IF(D4>40,D4-40,0)'
	cell_fill := sheet.rows[0].cells[0].fill or { return error('fill should be set') }
	assert cell_fill.theme == 5
	assert cell_fill.tint == 0.6
}

fn test_sheet_set_formula_currency_with_fill() ! {
	mut doc := Document.new()
	sheet_id := doc.add_sheet('Test')
	mut sheet := doc.get_sheet_mut(sheet_id)?

	loc := Location.from_encoding('N4')!
	fill := ThemeFill{
		theme: 9
		tint:  0.6
	}
	sheet.set_formula_currency_with_fill(loc, r'$C4*D4', .gbp, fill)

	assert sheet.rows.len == 1
	assert sheet.rows[0].cells.len == 1
	assert sheet.rows[0].cells[0].formula == r'$C4*D4'
	assert sheet.rows[0].cells[0].currency? == Currency.gbp
	cell_fill := sheet.rows[0].cells[0].fill or { return error('fill should be set') }
	assert cell_fill.theme == 9
	assert cell_fill.tint == 0.6
}

fn test_formula_with_absolute_reference() ! {
	mut doc := Document.new()
	sheet_id := doc.add_sheet('Test')
	mut sheet := doc.get_sheet_mut(sheet_id)?

	// Test that $ character is preserved in formulas (use raw string r'')
	sheet.set_formula(Location.from_encoding('N4')!, r'$C4*D4')

	assert sheet.rows[0].cells[0].formula == r'$C4*D4', 'absolute reference should be preserved'
}
