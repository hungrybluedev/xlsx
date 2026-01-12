module xlsx

import time

// build_cell creates a cell at the given location using the builder configuration.
// This is the primary method for creating styled cells with any combination of options.
//
// Example usage:
// ```v
// // Number with fill
// sheet.build_cell(loc, number: 42, fill: my_fill)
//
// // Formula with currency and fill
// sheet.build_cell(loc, formula: 'SUM(A1:A10)', currency: .gbp, fill: header_fill)
//
// // Text with fill
// sheet.build_cell(loc, text: 'Header', fill: highlight)
//
// // Date with fill
// sheet.build_cell(loc, date: my_date, fill: alt_row)
// ```
@[params]
pub fn (mut sheet Sheet) build_cell(loc Location, opts CellBuilder) {
	sheet.add_cell_internal(loc, opts)
}

// Internal implementation for all cell creation. Single source of truth.
fn (mut sheet Sheet) add_cell_internal(loc Location, opts CellBuilder) {
	sheet.ensure_row_exists(loc.row)
	row_idx := sheet.find_row_index(loc.row)

	// Determine cell type and value based on which content field is set
	mut cell_type := CellType.string_type
	mut value := ''
	mut formula_str := ''
	mut effective_style_id := opts.style_id

	if f := opts.formula {
		cell_type = .number_type
		formula_str = f
		value = ''
	} else if d := opts.date {
		cell_type = .number_type
		value = time_to_excel_date(d).str()
		if effective_style_id == 0 {
			effective_style_id = style_id_date
		}
	} else if n := opts.number {
		cell_type = .number_type
		// Format whole numbers without decimal places (e.g., 39 not 39.0)
		if n == f64(int(n)) {
			value = int(n).str()
		} else {
			value = n.str()
		}
	} else if t := opts.text {
		cell_type = .string_type
		value = t
	}

	sheet.rows[row_idx].cells << Cell{
		cell_type: cell_type
		location:  loc
		value:     value
		formula:   formula_str
		style_id:  effective_style_id
		currency:  opts.currency
		fill:      opts.fill
	}
}

// Sets a string cell at the given location
pub fn (mut sheet Sheet) set_cell(loc Location, value string) {
	sheet.add_cell_internal(loc, CellBuilder{ text: value })
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
	sheet.add_cell_internal(loc, CellBuilder{ number: value })
}

// Sets a formula cell at the given location
pub fn (mut sheet Sheet) set_formula(loc Location, formula string) {
	sheet.add_cell_internal(loc, CellBuilder{ formula: formula })
}

// Sets a date cell at the given location using a time.Time value.
// The time will be converted to Excel's serial date format internally.
// style_id=1 applies date formatting (e.g., "01-Jan")
pub fn (mut sheet Sheet) set_date(loc Location, date time.Time) {
	sheet.add_cell_internal(loc, CellBuilder{ date: date })
}

// Sets a currency cell at the given location with the specified currency format
pub fn (mut sheet Sheet) set_currency(loc Location, value f64, currency Currency) {
	sheet.add_cell_internal(loc, CellBuilder{ number: value, currency: currency })
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
