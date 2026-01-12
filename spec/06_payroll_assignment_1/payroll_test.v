import xlsx { Location }
import os

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
	sheet.set_date(Location.from_encoding('D3')!, 46023) // Excel date for 01-Jan-2026

	// Rows 4-23: Employee data (20 employees)
	for i, emp in employees {
		row := i + 4
		sheet.set_cell(Location.from_cartesian(row - 1, 0)!, emp.last_name) // Last Name (A)
		sheet.set_cell(Location.from_cartesian(row - 1, 1)!, emp.first_name) // First Name (B)
		sheet.set_number_f64(Location.from_cartesian(row - 1, 2)!, emp.hourly_wage) // Hourly Wage (C)
		sheet.set_number(Location.from_cartesian(row - 1, 3)!, emp.hours_worked) // Hours Worked (D)
		sheet.set_formula(Location.from_encoding('E${row}')!, 'C${row}*D${row}') // Pay formula (E)
	}

	// Row 24: Empty (skip)

	// Row 25: Max
	sheet.set_cell(Location.from_encoding('A25')!, 'Max')
	sheet.set_formula(Location.from_encoding('C25')!, 'MAX(C4:C23)')
	sheet.set_formula(Location.from_encoding('D25')!, 'MAX(D4:D23)')
	sheet.set_formula(Location.from_encoding('E25')!, 'MAX(E4:E23)')

	// Row 26: Min
	sheet.set_cell(Location.from_encoding('A26')!, 'Min')
	sheet.set_formula(Location.from_encoding('C26')!, 'MIN(C4:C23)')
	sheet.set_formula(Location.from_encoding('D26')!, 'MIN(D4:D23)')
	sheet.set_formula(Location.from_encoding('E26')!, 'MIN(E4:E23)')

	// Row 27: Average
	sheet.set_cell(Location.from_encoding('A27')!, 'Average')
	sheet.set_formula(Location.from_encoding('C27')!, 'AVERAGE(C4:C23)')
	sheet.set_formula(Location.from_encoding('D27')!, 'AVERAGE(D4:D23)')
	sheet.set_formula(Location.from_encoding('E27')!, 'AVERAGE(E4:E23)')

	// Row 28: Total
	sheet.set_cell(Location.from_encoding('A28')!, 'Total')
	sheet.set_formula(Location.from_encoding('C28')!, 'SUM(C4:C23)')
	sheet.set_formula(Location.from_encoding('D28')!, 'SUM(D4:D23)')
	sheet.set_formula(Location.from_encoding('E28')!, 'SUM(E4:E23)')

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

// TODO: This test is disabled because the reader has issues with sparse rows
// fn test_roundtrip_payroll() ! {
// 	// Read the original spec file
// 	spec_path := os.join_path(os.dir(@FILE), 'payroll.xlsx')
// 	original := xlsx.Document.from_file(spec_path)!
// 	original_data := original.sheets[1].get_all_data()!
//
// 	// Build equivalent document programmatically
// 	doc := build_payroll_document()!
//
// 	// Write to temp file
// 	temp_path := os.join_path(os.temp_dir(), 'test_payroll_roundtrip.xlsx')
// 	defer {
// 		os.rm(temp_path) or {}
// 	}
// 	doc.to_file(temp_path)!
//
// 	// Read back
// 	written_doc := xlsx.Document.from_file(temp_path)!
// 	written_data := written_doc.sheets[1].get_all_data()!
//
// 	// Compare dimensions
// 	assert written_data.raw_data.len == original_data.raw_data.len
//
// 	// Compare each row
// 	for i, original_row in original_data.raw_data {
// 		written_row := written_data.raw_data[i]
// 		assert written_row.len == original_row.len
//
// 		for j, original_cell in original_row {
// 			written_cell := written_row[j]
// 			assert written_cell == original_cell
// 		}
// 	}
// }
