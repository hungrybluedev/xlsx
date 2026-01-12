import xlsx { Location }
import os
import time
import math

// Helper function to compare cell values, handling floating point precision
fn values_match(ref_val string, out_val string) bool {
	// Direct string match
	if ref_val == out_val {
		return true
	}
	// Try parsing as floats and compare with tolerance
	ref_f := ref_val.f64()
	out_f := out_val.f64()
	// If both parse as valid numbers (not 0 when non-zero string), compare with tolerance
	if ref_f != 0.0 || ref_val == '0' || ref_val == '0.0' {
		if out_f != 0.0 || out_val == '0' || out_val == '0.0' {
			return math.abs(ref_f - out_f) < 1e-9
		}
	}
	return false
}

// Employee data structure
struct Employee {
	last_name    string
	first_name   string
	hourly_wage  f64
	hours_worked int
}

// Employee data from payroll.xlsx spec file
const employees = [
	Employee{'Rowntree', 'Geoffrey', 17.80, 39},
	Employee{'Sinclair', 'Felicity', 19.90, 36},
	Employee{'Thornton', 'Nigel', 16.45, 43},
	Employee{'Blackwood', 'Oliver', 16.20, 42},
	Employee{'Ashworth', 'Eleanor', 14.50, 38},
	Employee{'Hartley', 'Sebastian', 22.50, 45},
	Employee{'Ogilvie', 'Rosalind', 21.20, 39},
	Employee{'Pemberton', 'Hugh', 15.40, 37},
	Employee{'Quigley', 'Arabella', 14.25, 41},
	Employee{'Ingham', 'Cordelia', 13.25, 39},
	Employee{'Nettleton', 'Clive', 13.50, 42},
	Employee{'Cholmondeley', 'Harriet', 13.80, 35},
	Employee{'Darcy', 'Edmund', 19.40, 44},
	Employee{'Ellingham', 'Beatrice', 15.75, 40},
	Employee{'Jarvis', 'Theodore', 15.90, 43},
	Employee{'Moorhouse', 'Penelope', 16.70, 38},
	Employee{'Fairfax', 'Rupert', 17.30, 37},
	Employee{'Kensington', 'Imogen', 18.60, 36},
	Employee{'Langley', 'Alistair', 14.85, 40},
	Employee{'Grimshaw', 'Philippa', 14.10, 41},
]

fn build_payroll_document() !xlsx.Document {
	mut doc := xlsx.Document.new()
	sheet_id := doc.add_sheet('Sheet1')
	mut sheet := doc.get_sheet_mut(sheet_id) or { return error('Failed to get sheet') }

	// Row 1: Header
	sheet.set_cell(Location.from_encoding('A1')!, 'Employee Payroll')
	sheet.set_cell(Location.from_encoding('C1')!, 'Subhomoy Haldar')

	// Row 2: Column headers (partial)
	sheet.set_cell(Location.from_encoding('D2')!, 'Hours Worked')
	sheet.set_cell(Location.from_encoding('E2')!, 'Pay')

	// Row 3: Column headers
	sheet.set_cell(Location.from_encoding('A3')!, 'Last Name')
	sheet.set_cell(Location.from_encoding('B3')!, 'First Name')
	sheet.set_cell(Location.from_encoding('C3')!, 'Hourly Wage')
	// D3 contains a date (01-Jan) - use set_date for proper formatting
	jan_1_2026 := time.Time{
		year:  2026
		month: 1
		day:   1
	}
	sheet.set_date(Location.from_encoding('D3')!, jan_1_2026)

	// Rows 4-23: Employee data (20 employees)
	for i, emp in employees {
		row := i + 4
		sheet.set_cell(Location.from_cartesian(row - 1, 0)!, emp.last_name) // Last Name (A)
		sheet.set_cell(Location.from_cartesian(row - 1, 1)!, emp.first_name) // First Name (B)
		sheet.set_currency(Location.from_cartesian(row - 1, 2)!, emp.hourly_wage, .gbp) // Hourly Wage (C) with GBP
		sheet.set_number(Location.from_cartesian(row - 1, 3)!, emp.hours_worked) // Hours Worked (D)
		sheet.build_cell(Location.from_encoding('E${row}')!,
			formula:  'C${row}*D${row}'
			currency: .gbp
		) // Pay formula (E) with GBP
	}

	// Row 24: Empty (skip)

	// Row 25: Max
	sheet.set_cell(Location.from_encoding('A25')!, 'Max')
	sheet.build_cell(Location.from_encoding('C25')!, formula: 'MAX(C4:C23)', currency: .gbp)
	sheet.set_formula(Location.from_encoding('D25')!, 'MAX(D4:D23)')
	sheet.build_cell(Location.from_encoding('E25')!, formula: 'MAX(E4:E23)', currency: .gbp)

	// Row 26: Min
	sheet.set_cell(Location.from_encoding('A26')!, 'Min')
	sheet.build_cell(Location.from_encoding('C26')!, formula: 'MIN(C4:C23)', currency: .gbp)
	sheet.set_formula(Location.from_encoding('D26')!, 'MIN(D4:D23)')
	sheet.build_cell(Location.from_encoding('E26')!, formula: 'MIN(E4:E23)', currency: .gbp)

	// Row 27: Average
	sheet.set_cell(Location.from_encoding('A27')!, 'Average')
	sheet.build_cell(Location.from_encoding('C27')!, formula: 'AVERAGE(C4:C23)', currency: .gbp)
	sheet.set_formula(Location.from_encoding('D27')!, 'AVERAGE(D4:D23)')
	sheet.build_cell(Location.from_encoding('E27')!, formula: 'AVERAGE(E4:E23)', currency: .gbp)

	// Row 28: Total
	sheet.set_cell(Location.from_encoding('A28')!, 'Total')
	sheet.build_cell(Location.from_encoding('C28')!, formula: 'SUM(C4:C23)', currency: .gbp)
	sheet.set_formula(Location.from_encoding('D28')!, 'SUM(D4:D23)')
	sheet.build_cell(Location.from_encoding('E28')!, formula: 'SUM(E4:E23)', currency: .gbp)

	return doc
}

fn test_write_payroll() ! {
	// Build the payroll document programmatically
	doc := build_payroll_document()!

	// Write to a file in the spec directory for manual verification
	output_path := os.join_path(os.dir(@FILE), 'payroll_output.xlsx')
	doc.to_file(output_path)!

	// Verify file exists and is valid ZIP
	assert os.exists(output_path), 'output file should exist'
	content := os.read_bytes(output_path) or { return error('failed to read file') }
	assert content.len > 4, 'file should have content'
	assert content[0] == 0x50, 'should start with P (ZIP signature)'
	assert content[1] == 0x4B, 'should have K (ZIP signature)'

	// Note: Output file is kept at payroll_output.xlsx for manual verification
}

fn test_roundtrip_payroll() ! {
	// Build document programmatically
	doc := build_payroll_document()!

	// Write to temp file
	temp_path := os.join_path(os.temp_dir(), 'test_payroll_roundtrip.xlsx')
	defer {
		os.rm(temp_path) or {}
	}
	doc.to_file(temp_path)!

	// Read back
	written_doc := xlsx.Document.from_file(temp_path)!
	written_sheet := written_doc.sheets[1]
	written_data := written_sheet.get_all_data()!

	// Verify dimensions (28 rows, 5 columns: A-E)
	assert written_data.raw_data.len == 28, 'should have 28 rows'
	assert written_data.raw_data[0].len == 5, 'should have 5 columns'

	// Verify specific cells
	// Row 1: Headers
	assert written_data.raw_data[0][0] == 'Employee Payroll', 'A1 should be Employee Payroll'
	assert written_data.raw_data[0][2] == 'Subhomoy Haldar', 'C1 should be author name'

	// Row 3: Column headers
	assert written_data.raw_data[2][0] == 'Last Name', 'A3 should be Last Name'
	assert written_data.raw_data[2][1] == 'First Name', 'B3 should be First Name'
	assert written_data.raw_data[2][2] == 'Hourly Wage', 'C3 should be Hourly Wage'

	// Row 4: First employee data
	assert written_data.raw_data[3][0] == 'Rowntree', 'A4 should be Rowntree'
	assert written_data.raw_data[3][1] == 'Geoffrey', 'B4 should be Geoffrey'
	assert written_data.raw_data[3][2] == '17.8', 'C4 should be 17.8'
	assert written_data.raw_data[3][3] == '39', 'D4 should be 39'
	// E4 is a formula - value will be 0 (placeholder) until Excel recalculates

	// Row 24: Empty (sparse row)
	assert written_data.raw_data[23][0] == '', 'A24 should be empty'
	assert written_data.raw_data[23][4] == '', 'E24 should be empty'

	// Row 25: Summary (Max)
	assert written_data.raw_data[24][0] == 'Max', 'A25 should be Max'

	// Row 28: Summary (Total)
	assert written_data.raw_data[27][0] == 'Total', 'A28 should be Total'
}

fn test_compare_reference_vs_generated() ! {
	// Read reference file (source of truth)
	ref_path := os.join_path(os.dir(@FILE), 'payroll.xlsx')
	ref_doc := xlsx.Document.from_file(ref_path)!
	ref_sheet := ref_doc.sheets[1]

	// Build and write output file
	doc := build_payroll_document()!
	output_path := os.join_path(os.dir(@FILE), 'payroll_output.xlsx')
	doc.to_file(output_path)!

	// Read generated file
	out_doc := xlsx.Document.from_file(output_path)!
	out_sheet := out_doc.sheets[1]

	// Compare dimensions
	assert ref_sheet.top_left.row == out_sheet.top_left.row, 'top_left row mismatch'
	assert ref_sheet.top_left.col == out_sheet.top_left.col, 'top_left col mismatch'
	assert ref_sheet.bottom_right.row == out_sheet.bottom_right.row, 'bottom_right row mismatch'
	assert ref_sheet.bottom_right.col == out_sheet.bottom_right.col, 'bottom_right col mismatch'

	// Compare all cell values (excluding formula cells which have placeholder values in generated files)
	ref_data := ref_sheet.get_all_data()!
	out_data := out_sheet.get_all_data()!

	assert ref_data.raw_data.len == out_data.raw_data.len, 'row count mismatch: ref=${ref_data.raw_data.len}, out=${out_data.raw_data.len}'

	for row_idx, ref_row in ref_data.raw_data {
		assert ref_row.len == out_data.raw_data[row_idx].len, 'col count mismatch at row ${row_idx}'
		for col_idx, ref_val in ref_row {
			out_val := out_data.raw_data[row_idx][col_idx]
			// Check if this is a formula cell
			// Column E (index 4) has formulas in rows 4-23 and 25-28
			// Columns C, D, E (indices 2, 3, 4) have formulas in summary rows 25-28
			is_summary_row := row_idx >= 24 && row_idx <= 27
			is_formula_cell := (col_idx == 4 && row_idx >= 3) || (is_summary_row && col_idx >= 2)
			if is_formula_cell {
				// For formula cells, the generated file has placeholder 0, so skip value comparison
				// Formula comparison is done separately below
				continue
			}
			// Compare values - handle floating point precision differences
			if !values_match(ref_val, out_val) {
				assert false, 'value mismatch at row ${row_idx + 1}, col ${col_idx}: expected "${ref_val}", got "${out_val}"'
			}
		}
	}

	// Compare formulas for cells with formulas
	// E4 should have formula C4*D4 (not a shared formula, should match exactly)
	e4_ref := ref_sheet.get_cell(Location.from_encoding('E4')!) or {
		assert false, 'E4 not found in reference'
		return
	}
	e4_out := out_sheet.get_cell(Location.from_encoding('E4')!) or {
		assert false, 'E4 not found in output'
		return
	}
	assert e4_ref.formula == e4_out.formula, 'E4 formula mismatch: ref="${e4_ref.formula}", out="${e4_out.formula}"'

	// E25 should have a MAX formula
	// Note: shared formulas in ref file use base cell references, so we check both have formulas
	e25_ref := ref_sheet.get_cell(Location.from_encoding('E25')!) or {
		assert false, 'E25 not found in reference'
		return
	}
	e25_out := out_sheet.get_cell(Location.from_encoding('E25')!) or {
		assert false, 'E25 not found in output'
		return
	}
	// Both should have formulas containing MAX
	assert e25_ref.formula.contains('MAX'), 'E25 ref should have MAX formula, got: ${e25_ref.formula}'
	assert e25_out.formula.contains('MAX'), 'E25 out should have MAX formula, got: ${e25_out.formula}'

	// Compare currency formatting
	// C4 should have GBP currency
	c4_ref := ref_sheet.get_cell(Location.from_encoding('C4')!) or {
		assert false, 'C4 not found in reference'
		return
	}
	c4_out := out_sheet.get_cell(Location.from_encoding('C4')!) or {
		assert false, 'C4 not found in output'
		return
	}
	assert c4_ref.currency == c4_out.currency, 'C4 currency mismatch: ref=${c4_ref.currency}, out=${c4_out.currency}'
}
