module xlsx

// CellOptions provides a unified way to configure cell properties.
// Use this with set_cell_with_options for complex cell configurations
// instead of the many specialized setter methods.
pub struct CellOptions {
pub:
	fill       ?ThemeFill // Background fill color
	currency   ?Currency  // Currency formatting
	style_id   int        // Style ID (use style_id_date for dates)
	is_formula bool       // Whether value is a formula
}

// Sets a cell with flexible options at the given location.
// This is a unified method that can replace most other setters.
//
// Example usage:
// ```v
// // Simple string cell
// sheet.set_cell_with_options(loc, 'Hello', CellOptions{})
//
// // Number with fill
// sheet.set_cell_with_options(loc, '42', CellOptions{ fill: my_fill })
//
// // Formula with currency and fill
// sheet.set_cell_with_options(loc, 'SUM(A1:A10)', CellOptions{
//     is_formula: true
//     currency: Currency{ symbol: '$', decimal_places: 2 }
//     fill: my_fill
// })
// ```
pub fn (mut sheet Sheet) set_cell_with_options(loc Location, value string, opts CellOptions) {
	sheet.ensure_row_exists(loc.row)
	row_idx := sheet.find_row_index(loc.row)

	cell_type := if opts.is_formula { CellType.number_type } else { CellType.string_type }
	formula := if opts.is_formula { value } else { '' }
	cell_value := if opts.is_formula { '' } else { value }

	sheet.rows[row_idx].cells << Cell{
		cell_type: cell_type
		location:  loc
		value:     cell_value
		formula:   formula
		style_id:  opts.style_id
		currency:  opts.currency
		fill:      opts.fill
	}
}

// Sets a numeric cell with flexible options at the given location.
pub fn (mut sheet Sheet) set_number_with_options(loc Location, value f64, opts CellOptions) {
	sheet.ensure_row_exists(loc.row)
	row_idx := sheet.find_row_index(loc.row)

	sheet.rows[row_idx].cells << Cell{
		cell_type: .number_type
		location:  loc
		value:     value.str()
		style_id:  opts.style_id
		currency:  opts.currency
		fill:      opts.fill
	}
}

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
		style_id:  style_id_date
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

// Sets a numeric cell with a background fill color
pub fn (mut sheet Sheet) set_number_with_fill(loc Location, value int, fill ThemeFill) {
	sheet.ensure_row_exists(loc.row)
	row_idx := sheet.find_row_index(loc.row)
	sheet.rows[row_idx].cells << Cell{
		cell_type: .number_type
		location:  loc
		value:     value.str()
		fill:      fill
	}
}

// Sets a date cell with a background fill color
pub fn (mut sheet Sheet) set_date_with_fill(loc Location, excel_date int, fill ThemeFill) {
	sheet.ensure_row_exists(loc.row)
	row_idx := sheet.find_row_index(loc.row)
	sheet.rows[row_idx].cells << Cell{
		cell_type: .number_type
		location:  loc
		value:     excel_date.str()
		style_id:  style_id_date
		fill:      fill
	}
}

// Sets a formula cell with a background fill color (no currency)
pub fn (mut sheet Sheet) set_formula_with_fill(loc Location, formula string, fill ThemeFill) {
	sheet.ensure_row_exists(loc.row)
	row_idx := sheet.find_row_index(loc.row)
	sheet.rows[row_idx].cells << Cell{
		cell_type: .number_type
		location:  loc
		value:     '' // Value will be computed by Excel
		formula:   formula
		fill:      fill
	}
}

// Sets a formula cell with currency formatting and a background fill color
pub fn (mut sheet Sheet) set_formula_currency_with_fill(loc Location, formula string, currency Currency, fill ThemeFill) {
	sheet.ensure_row_exists(loc.row)
	row_idx := sheet.find_row_index(loc.row)
	sheet.rows[row_idx].cells << Cell{
		cell_type: .number_type
		location:  loc
		value:     '' // Value will be computed by Excel
		formula:   formula
		currency:  currency
		fill:      fill
	}
}
