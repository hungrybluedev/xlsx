module xlsx

import strings

const a_ascii = u8(`A`)
const max_rows = 1048576
const max_cols = 16384

fn col_to_label(col int) string {
	if col < 26 {
		col_char := u8(col) + a_ascii
		return col_char.ascii_str()
	}
	return col_to_label(col / 26 - 1) + col_to_label(col % 26)
}

fn label_to_col(label string) int {
	mut col := 0
	for ch in label {
		col *= 26
		col += ch - a_ascii + 1
	}
	return col - 1
}

pub fn Location.from_cartesian(row int, col int) !Location {
	if row < 0 {
		return error('Row must be >= 0')
	}
	if row > max_rows {
		return error('Row must be <= ${max_rows}')
	}
	if col < 0 {
		return error('Col must be >= 0')
	}
	if col > max_cols {
		return error('Col must be <= ${max_cols}')
	}

	return Location{
		row:       row
		col:       col
		row_label: (row + 1).str()
		col_label: col_to_label(col)
	}
}

pub fn Location.from_encoding(code string) !Location {
	if code.len < 2 {
		return error('Invalid location code. Must be at least 2 characters long.')
	}

	mut column_buffer := strings.new_builder(8)
	mut row_buffer := strings.new_builder(8)

	for location, ch in code {
		if ch.is_digit() {
			row_buffer.write_string(code[location..])
			break
		}
		column_buffer.write_u8(ch)
	}

	row := row_buffer.str()
	col := column_buffer.str()

	return Location{
		row:       row.int() - 1
		col:       label_to_col(col)
		row_label: row
		col_label: col
	}
}
