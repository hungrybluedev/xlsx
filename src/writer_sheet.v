module xlsx

// Sets a string cell at the given location
pub fn (mut sheet Sheet) set_cell(loc Location, value string) {
	sheet.ensure_row_exists(loc.row)
	row_idx := sheet.find_row_index(loc.row)
	sheet.rows[row_idx].cells << Cell{
		cell_type: .string_type
		location:  loc
		value:     value
	}
}

// Sets a numeric cell at the given location (int version)
pub fn (mut sheet Sheet) set_number(loc Location, value int) {
	sheet.ensure_row_exists(loc.row)
	row_idx := sheet.find_row_index(loc.row)
	sheet.rows[row_idx].cells << Cell{
		cell_type: .number_type
		location:  loc
		value:     value.str()
	}
}

// Sets a numeric cell at the given location (f64 version)
pub fn (mut sheet Sheet) set_number_f64(loc Location, value f64) {
	sheet.ensure_row_exists(loc.row)
	row_idx := sheet.find_row_index(loc.row)
	sheet.rows[row_idx].cells << Cell{
		cell_type: .number_type
		location:  loc
		value:     value.str()
	}
}

// Sets a formula cell at the given location
pub fn (mut sheet Sheet) set_formula(loc Location, formula string) {
	sheet.ensure_row_exists(loc.row)
	row_idx := sheet.find_row_index(loc.row)
	sheet.rows[row_idx].cells << Cell{
		cell_type: .number_type // Formulas are typically numbers
		location:  loc
		value:     '' // Value will be computed by Excel
		formula:   formula
	}
}

// Sets a date cell at the given location (Excel serial date number)
// Excel dates are stored as numbers: days since 1900-01-01 (with 1900 bug)
// style_id=1 applies date formatting (e.g., "01-Jan")
pub fn (mut sheet Sheet) set_date(loc Location, excel_date int) {
	sheet.ensure_row_exists(loc.row)
	row_idx := sheet.find_row_index(loc.row)
	sheet.rows[row_idx].cells << Cell{
		cell_type: .number_type
		location:  loc
		value:     excel_date.str()
		style_id:  1 // Date format style
	}
}

// Ensures a row exists at the given index, creating it if necessary
fn (mut sheet Sheet) ensure_row_exists(row_index int) {
	// Check if row already exists
	for r in sheet.rows {
		if r.row_index == row_index {
			return
		}
	}
	// Create new row
	sheet.rows << Row{
		row_index: row_index
		row_label: (row_index + 1).str()
		cells:     []Cell{}
	}
}

// Finds the internal index of a row by its row_index
fn (sheet Sheet) find_row_index(row_index int) int {
	for i, r in sheet.rows {
		if r.row_index == row_index {
			return i
		}
	}
	return -1 // Should not happen if ensure_row_exists was called
}

// Sets a currency cell at the given location with the specified currency format
// The currency format will be applied when the file is written
pub fn (mut sheet Sheet) set_currency(loc Location, value f64, currency Currency) {
	sheet.ensure_row_exists(loc.row)
	row_idx := sheet.find_row_index(loc.row)
	sheet.rows[row_idx].cells << Cell{
		cell_type: .number_type
		location:  loc
		value:     value.str()
		currency:  currency
	}
}

// Sets a formula cell with currency formatting
pub fn (mut sheet Sheet) set_formula_currency(loc Location, formula string, currency Currency) {
	sheet.ensure_row_exists(loc.row)
	row_idx := sheet.find_row_index(loc.row)
	sheet.rows[row_idx].cells << Cell{
		cell_type: .number_type
		location:  loc
		value:     '' // Value will be computed by Excel
		formula:   formula
		currency:  currency
	}
}

// Sets a formula cell with a specific style at the given location
// Kept for backward compatibility with non-currency styles
pub fn (mut sheet Sheet) set_formula_with_style(loc Location, formula string, style_id int) {
	sheet.ensure_row_exists(loc.row)
	row_idx := sheet.find_row_index(loc.row)
	sheet.rows[row_idx].cells << Cell{
		cell_type: .number_type
		location:  loc
		value:     '' // Value will be computed by Excel
		formula:   formula
		style_id:  style_id
	}
}
