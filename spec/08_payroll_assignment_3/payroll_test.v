import xlsx { Location, ThemeFill }
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

// Helper function to compare floating point values with tolerance
fn f64_approx_equal(a f64, b f64, tolerance f64) bool {
	return math.abs(a - b) < tolerance
}

// Employee data structure with 5 weeks of hours
struct EmployeeWeekly {
	last_name    string
	first_name   string
	hourly_wage  f64
	hours_worked [5]int // 5 weeks: 01-Jan, 08-Jan, 15-Jan, 22-Jan, 29-Jan
}

// Fill colors for each column section (from spec/08 expected file)
// Theme indices map to Office theme colors defined in theme1.xml:
// Theme 3 = dk2 (0E2841 dark blue), Theme 5 = accent2 (E97132 orange)
// Theme 8 = accent5 (A02B93 purple), Theme 9 = accent6 (4EA72E green)
const fill_hours_worked = ThemeFill{
	theme: 3
	tint:  0.749992370372631
} // Light blue (D-H)
const fill_overtime_hours = ThemeFill{
	theme: 5
	tint:  0.59999389629810485
} // Light orange (I-M)
const fill_pay = ThemeFill{
	theme: 9
	tint:  0.79998168889431442
} // Very light green (S-W)
const fill_overtime_bonus = ThemeFill{
	theme: 8
	tint:  0.79998168889431442
} // Light purple (X-AB)
const fill_total = ThemeFill{
	theme: 9
	tint:  0.59999389629810485
} // Light green (N-R)

// Employee data with 5 weeks of hours (same employees as spec/06 and spec/07)
// Hours extracted from spec/08 expected file
const employees = [
	EmployeeWeekly{'Rowntree', 'Geoffrey', 17.80, [39, 36, 40, 41, 38]!},
	EmployeeWeekly{'Sinclair', 'Felicity', 19.90, [36, 42, 35, 39, 41]!},
	EmployeeWeekly{'Thornton', 'Nigel', 16.45, [43, 41, 40, 39, 40]!},
	EmployeeWeekly{'Blackwood', 'Oliver', 16.20, [42, 40, 41, 39, 39]!},
	EmployeeWeekly{'Ashworth', 'Eleanor', 14.50, [38, 40, 41, 38, 38]!},
	EmployeeWeekly{'Hartley', 'Sebastian', 22.50, [45, 42, 41, 39, 41]!},
	EmployeeWeekly{'Ogilvie', 'Rosalind', 21.20, [39, 37, 39, 38, 37]!},
	EmployeeWeekly{'Pemberton', 'Hugh', 15.40, [37, 39, 36, 34, 39]!},
	EmployeeWeekly{'Quigley', 'Arabella', 14.25, [41, 39, 40, 38, 37]!},
	EmployeeWeekly{'Ingham', 'Cordelia', 13.25, [39, 39, 39, 40, 39]!},
	EmployeeWeekly{'Nettleton', 'Clive', 13.50, [42, 40, 41, 40, 41]!},
	EmployeeWeekly{'Cholmondeley', 'Harriet', 13.80, [35, 38, 39, 35, 39]!},
	EmployeeWeekly{'Darcy', 'Edmund', 19.40, [44, 45, 40, 42, 45]!},
	EmployeeWeekly{'Ellingham', 'Beatrice', 15.75, [40, 40, 39, 40, 39]!},
	EmployeeWeekly{'Jarvis', 'Theodore', 15.90, [43, 42, 41, 39, 40]!},
	EmployeeWeekly{'Moorhouse', 'Penelope', 16.70, [38, 39, 40, 42, 40]!},
	EmployeeWeekly{'Fairfax', 'Rupert', 17.30, [37, 38, 37, 35, 40]!},
	EmployeeWeekly{'Kensington', 'Imogen', 18.60, [36, 37, 38, 35, 38]!},
	EmployeeWeekly{'Langley', 'Alistair', 14.85, [40, 40, 40, 41, 37]!},
	EmployeeWeekly{'Grimshaw', 'Philippa', 14.10, [41, 40, 42, 39, 41]!},
]

// Column labels for each week (D-H, I-M, N-R, S-W, X-AB)
const hours_cols = ['D', 'E', 'F', 'G', 'H']
const overtime_cols = ['I', 'J', 'K', 'L', 'M']
const pay_cols = ['N', 'O', 'P', 'Q', 'R']
const bonus_cols = ['S', 'T', 'U', 'V', 'W']
const total_cols = ['X', 'Y', 'Z', 'AA', 'AB']

fn build_payroll_document() !xlsx.Document {
	mut doc := xlsx.Document.new()
	sheet_id := doc.add_sheet('Sheet1')
	mut sheet := doc.get_sheet_mut(sheet_id) or { return error('Failed to get sheet') }

	// Row 1: Header
	sheet.set_cell(Location.from_encoding('A1')!, 'Employee Payroll')
	sheet.set_cell(Location.from_encoding('C1')!, 'Subhomoy Haldar')

	// Row 2: Section headers
	sheet.set_cell(Location.from_encoding('D2')!, 'Hours Worked')
	sheet.set_cell(Location.from_encoding('I2')!, 'Overtime Hours')
	sheet.set_cell(Location.from_encoding('N2')!, 'Pay')
	sheet.set_cell(Location.from_encoding('S2')!, 'Overtime Bonus')
	sheet.set_cell(Location.from_encoding('X2')!, 'Total')
	sheet.set_cell(Location.from_encoding('AD2')!, 'January Pay')

	// Row 3: Column headers
	sheet.set_cell(Location.from_encoding('A3')!, 'Last Name')
	sheet.set_cell(Location.from_encoding('B3')!, 'First Name')
	sheet.set_cell(Location.from_encoding('C3')!, 'Hourly Wage')

	// Date headers for each week (01-Jan to 29-Jan, +7 days each)
	base_date := time.Time{
		year:  2026
		month: 1
		day:   1
	}

	// Hours Worked dates (D3-H3)
	for week in 0 .. 5 {
		col := hours_cols[week]
		week_date := base_date.add_days(week * 7)
		sheet.build_cell(Location.from_encoding('${col}3')!,
			date: week_date
			fill: fill_hours_worked
		)
	}

	// Overtime Hours dates (I3-M3)
	for week in 0 .. 5 {
		col := overtime_cols[week]
		week_date := base_date.add_days(week * 7)
		sheet.build_cell(Location.from_encoding('${col}3')!,
			date: week_date
			fill: fill_overtime_hours
		)
	}

	// Pay dates (N3-R3)
	for week in 0 .. 5 {
		col := pay_cols[week]
		week_date := base_date.add_days(week * 7)
		sheet.build_cell(Location.from_encoding('${col}3')!, date: week_date, fill: fill_pay)
	}

	// Overtime Bonus dates (S3-W3)
	for week in 0 .. 5 {
		col := bonus_cols[week]
		week_date := base_date.add_days(week * 7)
		sheet.build_cell(Location.from_encoding('${col}3')!,
			date: week_date
			fill: fill_overtime_bonus
		)
	}

	// Total dates (X3-AB3)
	for week in 0 .. 5 {
		col := total_cols[week]
		week_date := base_date.add_days(week * 7)
		sheet.build_cell(Location.from_encoding('${col}3')!, date: week_date, fill: fill_total)
	}

	// Rows 4-23: Employee data (20 employees)
	for i, emp in employees {
		row := i + 4

		// Basic info (A-C)
		sheet.set_cell(Location.from_cartesian(row - 1, 0)!, emp.last_name) // A
		sheet.set_cell(Location.from_cartesian(row - 1, 1)!, emp.first_name) // B
		sheet.set_currency(Location.from_cartesian(row - 1, 2)!, emp.hourly_wage, .gbp) // C

		// Hours Worked (D-H) with light blue fill
		for week in 0 .. 5 {
			col := hours_cols[week]
			sheet.build_cell(Location.from_encoding('${col}${row}')!,
				number: emp.hours_worked[week]
				fill:   fill_hours_worked
			)
		}

		// Overtime Hours (I-M) with light cyan fill - IF formula with absolute ref
		for week in 0 .. 5 {
			hours_col := hours_cols[week]
			ot_col := overtime_cols[week]
			// Formula: IF(D4>40,D4-40,0) etc.
			sheet.build_cell(Location.from_encoding('${ot_col}${row}')!,
				formula: 'IF(${hours_col}${row}>40,${hours_col}${row}-40,0)'
				fill:    fill_overtime_hours
			)
		}

		// Pay (N-R) with orange fill - absolute reference to hourly wage
		for week in 0 .. 5 {
			hours_col := hours_cols[week]
			pay_col := pay_cols[week]
			// Formula: $C4*D4 (absolute reference to C column)
			sheet.build_cell(Location.from_encoding('${pay_col}${row}')!,
				formula:  r'$C' + '${row}*${hours_col}${row}'
				currency: .gbp
				fill:     fill_pay
			)
		}

		// Overtime Bonus (S-W) with light orange fill
		for week in 0 .. 5 {
			ot_col := overtime_cols[week]
			bonus_col := bonus_cols[week]
			// Formula: 0.5*$C4*I4 (50% of wage * overtime hours)
			sheet.build_cell(Location.from_encoding('${bonus_col}${row}')!,
				formula:  r'0.5*$C' + '${row}*${ot_col}${row}'
				currency: .gbp
				fill:     fill_overtime_bonus
			)
		}

		// Total (X-AB) with green fill
		for week in 0 .. 5 {
			pay_col := pay_cols[week]
			bonus_col := bonus_cols[week]
			total_col := total_cols[week]
			// Formula: N4+S4 (pay + bonus)
			sheet.build_cell(Location.from_encoding('${total_col}${row}')!,
				formula:  '${pay_col}${row}+${bonus_col}${row}'
				currency: .gbp
				fill:     fill_total
			)
		}

		// January Pay (AD) - sum of all weekly totals
		sheet.build_cell(Location.from_encoding('AD${row}')!,
			formula:  'SUM(X${row}:AB${row})'
			currency: .gbp
		)
	}

	// Row 24: Empty (skip)

	// Row 25: Max
	sheet.set_cell(Location.from_encoding('A25')!, 'Max')
	sheet.build_cell(Location.from_encoding('C25')!, formula: 'MAX(C4:C23)', currency: .gbp)
	for week in 0 .. 5 {
		hours_col := hours_cols[week]
		sheet.set_formula(Location.from_encoding('${hours_col}25')!, 'MAX(${hours_col}4:${hours_col}23)')
	}
	for week in 0 .. 5 {
		pay_col := pay_cols[week]
		sheet.build_cell(Location.from_encoding('${pay_col}25')!,
			formula:  'MAX(${pay_col}4:${pay_col}23)'
			currency: .gbp
		)
	}
	for week in 0 .. 5 {
		bonus_col := bonus_cols[week]
		sheet.build_cell(Location.from_encoding('${bonus_col}25')!,
			formula:  'MAX(${bonus_col}4:${bonus_col}23)'
			currency: .gbp
		)
	}
	for week in 0 .. 5 {
		total_col := total_cols[week]
		sheet.build_cell(Location.from_encoding('${total_col}25')!,
			formula:  'MAX(${total_col}4:${total_col}23)'
			currency: .gbp
		)
	}
	// January Pay Max (AD25)
	sheet.build_cell(Location.from_encoding('AD25')!, formula: 'MAX(AD4:AD23)', currency: .gbp)

	// Row 26: Min
	sheet.set_cell(Location.from_encoding('A26')!, 'Min')
	sheet.build_cell(Location.from_encoding('C26')!, formula: 'MIN(C4:C23)', currency: .gbp)
	for week in 0 .. 5 {
		hours_col := hours_cols[week]
		sheet.set_formula(Location.from_encoding('${hours_col}26')!, 'MIN(${hours_col}4:${hours_col}23)')
	}
	for week in 0 .. 5 {
		pay_col := pay_cols[week]
		sheet.build_cell(Location.from_encoding('${pay_col}26')!,
			formula:  'MIN(${pay_col}4:${pay_col}23)'
			currency: .gbp
		)
	}
	for week in 0 .. 5 {
		bonus_col := bonus_cols[week]
		sheet.build_cell(Location.from_encoding('${bonus_col}26')!,
			formula:  'MIN(${bonus_col}4:${bonus_col}23)'
			currency: .gbp
		)
	}
	for week in 0 .. 5 {
		total_col := total_cols[week]
		sheet.build_cell(Location.from_encoding('${total_col}26')!,
			formula:  'MIN(${total_col}4:${total_col}23)'
			currency: .gbp
		)
	}
	// January Pay Min (AD26)
	sheet.build_cell(Location.from_encoding('AD26')!, formula: 'MIN(AD4:AD23)', currency: .gbp)

	// Row 27: Average
	sheet.set_cell(Location.from_encoding('A27')!, 'Average')
	sheet.build_cell(Location.from_encoding('C27')!, formula: 'AVERAGE(C4:C23)', currency: .gbp)
	for week in 0 .. 5 {
		hours_col := hours_cols[week]
		sheet.set_formula(Location.from_encoding('${hours_col}27')!, 'AVERAGE(${hours_col}4:${hours_col}23)')
	}
	for week in 0 .. 5 {
		pay_col := pay_cols[week]
		sheet.build_cell(Location.from_encoding('${pay_col}27')!,
			formula:  'AVERAGE(${pay_col}4:${pay_col}23)'
			currency: .gbp
		)
	}
	for week in 0 .. 5 {
		bonus_col := bonus_cols[week]
		sheet.build_cell(Location.from_encoding('${bonus_col}27')!,
			formula:  'AVERAGE(${bonus_col}4:${bonus_col}23)'
			currency: .gbp
		)
	}
	for week in 0 .. 5 {
		total_col := total_cols[week]
		sheet.build_cell(Location.from_encoding('${total_col}27')!,
			formula:  'AVERAGE(${total_col}4:${total_col}23)'
			currency: .gbp
		)
	}
	// January Pay Average (AD27)
	sheet.build_cell(Location.from_encoding('AD27')!, formula: 'AVERAGE(AD4:AD23)', currency: .gbp)

	// Row 28: Total (Sum)
	sheet.set_cell(Location.from_encoding('A28')!, 'Total')
	sheet.build_cell(Location.from_encoding('C28')!, formula: 'SUM(C4:C23)', currency: .gbp)
	for week in 0 .. 5 {
		hours_col := hours_cols[week]
		sheet.set_formula(Location.from_encoding('${hours_col}28')!, 'SUM(${hours_col}4:${hours_col}23)')
	}
	for week in 0 .. 5 {
		pay_col := pay_cols[week]
		sheet.build_cell(Location.from_encoding('${pay_col}28')!,
			formula:  'SUM(${pay_col}4:${pay_col}23)'
			currency: .gbp
		)
	}
	for week in 0 .. 5 {
		bonus_col := bonus_cols[week]
		sheet.build_cell(Location.from_encoding('${bonus_col}28')!,
			formula:  'SUM(${bonus_col}4:${bonus_col}23)'
			currency: .gbp
		)
	}
	for week in 0 .. 5 {
		total_col := total_cols[week]
		sheet.build_cell(Location.from_encoding('${total_col}28')!,
			formula:  'SUM(${total_col}4:${total_col}23)'
			currency: .gbp
		)
	}
	// January Pay Total (AD28)
	sheet.build_cell(Location.from_encoding('AD28')!, formula: 'SUM(AD4:AD23)', currency: .gbp)

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
	temp_path := os.join_path(os.temp_dir(), 'test_payroll_spec08_roundtrip.xlsx')
	defer {
		os.rm(temp_path) or {}
	}
	doc.to_file(temp_path)!

	// Read back
	written_doc := xlsx.Document.from_file(temp_path)!
	written_sheet := written_doc.sheets[1]
	written_data := written_sheet.get_all_data()!

	// Verify dimensions (28 rows, 30 columns: A-AD)
	assert written_data.raw_data.len == 28, 'should have 28 rows, got ${written_data.raw_data.len}'
	assert written_data.raw_data[0].len == 30, 'should have 30 columns, got ${written_data.raw_data[0].len}'

	// Verify specific cells
	// Row 1: Headers
	assert written_data.raw_data[0][0] == 'Employee Payroll', 'A1 should be Employee Payroll'
	assert written_data.raw_data[0][2] == 'Subhomoy Haldar', 'C1 should be author name'

	// Row 2: Section headers
	assert written_data.raw_data[1][3] == 'Hours Worked', 'D2 should be Hours Worked'
	assert written_data.raw_data[1][8] == 'Overtime Hours', 'I2 should be Overtime Hours'
	assert written_data.raw_data[1][13] == 'Pay', 'N2 should be Pay'
	assert written_data.raw_data[1][18] == 'Overtime Bonus', 'S2 should be Overtime Bonus'
	assert written_data.raw_data[1][23] == 'Total', 'X2 should be Total'
	assert written_data.raw_data[1][29] == 'January Pay', 'AD2 should be January Pay'

	// Row 3: Column headers
	assert written_data.raw_data[2][0] == 'Last Name', 'A3 should be Last Name'
	assert written_data.raw_data[2][1] == 'First Name', 'B3 should be First Name'
	assert written_data.raw_data[2][2] == 'Hourly Wage', 'C3 should be Hourly Wage'

	// Row 4: First employee data
	assert written_data.raw_data[3][0] == 'Rowntree', 'A4 should be Rowntree'
	assert written_data.raw_data[3][1] == 'Geoffrey', 'B4 should be Geoffrey'
	assert written_data.raw_data[3][2] == '17.8', 'C4 should be 17.8'
	assert written_data.raw_data[3][3] == '39', 'D4 should be 39 (first week hours)'

	// Row 24: Empty (sparse row)
	assert written_data.raw_data[23][0] == '', 'A24 should be empty'

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
			// Data rows 4-23 (indices 3-22): cols I-AD (indices 8-29) have formulas
			// Summary rows 25-28 (indices 24-27): cols C onwards have formulas
			is_data_row := row_idx >= 3 && row_idx <= 22
			is_summary_row := row_idx >= 24 && row_idx <= 27
			is_formula_cell := (is_data_row && col_idx >= 8) || (is_summary_row && col_idx >= 2)
			if is_formula_cell {
				// For formula cells, the generated file has placeholder 0, so skip value comparison
				continue
			}
			// Compare values - handle floating point precision differences
			if !values_match(ref_val, out_val) {
				assert false, 'value mismatch at row ${row_idx + 1}, col ${col_idx}: expected "${ref_val}", got "${out_val}"'
			}
		}
	}

	// Compare formulas for key cells
	// D4 should have hours value (not a formula)
	d4_ref := ref_sheet.get_cell(Location.from_encoding('D4')!) or {
		assert false, 'D4 not found in reference'
		return
	}
	d4_out := out_sheet.get_cell(Location.from_encoding('D4')!) or {
		assert false, 'D4 not found in output'
		return
	}
	assert d4_ref.value == d4_out.value, 'D4 value mismatch: ref="${d4_ref.value}", out="${d4_out.value}"'

	// I4 should have IF formula for overtime hours
	i4_ref := ref_sheet.get_cell(Location.from_encoding('I4')!) or {
		assert false, 'I4 not found in reference'
		return
	}
	i4_out := out_sheet.get_cell(Location.from_encoding('I4')!) or {
		assert false, 'I4 not found in output'
		return
	}
	assert i4_ref.formula.contains('IF'), 'I4 ref should have IF formula, got: ${i4_ref.formula}'
	assert i4_out.formula.contains('IF'), 'I4 out should have IF formula, got: ${i4_out.formula}'

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

	// Compare theme fills
	// D4 should have fill_hours_worked (theme 3)
	assert d4_ref.fill != none, 'D4 ref should have fill'
	assert d4_out.fill != none, 'D4 out should have fill'
	if ref_fill := d4_ref.fill {
		if out_fill := d4_out.fill {
			assert ref_fill.theme == out_fill.theme, 'D4 fill theme mismatch: ref=${ref_fill.theme}, out=${out_fill.theme}'
			assert f64_approx_equal(ref_fill.tint, out_fill.tint, 1e-10), 'D4 fill tint mismatch: ref=${ref_fill.tint}, out=${out_fill.tint}'
		}
	}

	// I4 should have fill_overtime_hours (theme 5)
	assert i4_ref.fill != none, 'I4 ref should have fill'
	assert i4_out.fill != none, 'I4 out should have fill'
	if ref_fill := i4_ref.fill {
		if out_fill := i4_out.fill {
			assert ref_fill.theme == out_fill.theme, 'I4 fill theme mismatch: ref=${ref_fill.theme}, out=${out_fill.theme}'
			assert f64_approx_equal(ref_fill.tint, out_fill.tint, 1e-10), 'I4 fill tint mismatch: ref=${ref_fill.tint}, out=${out_fill.tint}'
		}
	}

	// N4 should have fill_pay (theme 9, lighter green)
	n4_ref := ref_sheet.get_cell(Location.from_encoding('N4')!) or {
		assert false, 'N4 not found in reference'
		return
	}
	n4_out := out_sheet.get_cell(Location.from_encoding('N4')!) or {
		assert false, 'N4 not found in output'
		return
	}
	assert n4_ref.fill != none, 'N4 ref should have fill'
	assert n4_out.fill != none, 'N4 out should have fill'
	if ref_fill := n4_ref.fill {
		if out_fill := n4_out.fill {
			assert ref_fill.theme == out_fill.theme, 'N4 fill theme mismatch: ref=${ref_fill.theme}, out=${out_fill.theme}'
			assert f64_approx_equal(ref_fill.tint, out_fill.tint, 1e-10), 'N4 fill tint mismatch: ref=${ref_fill.tint}, out=${out_fill.tint}'
		}
	}
}
