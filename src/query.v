module xlsx

pub fn (sheet Sheet) get_cell(location Location) ?Cell {
	if location.row >= sheet.rows.len {
		return none
	}
	target_row := sheet.rows[location.row]
	if location.col >= target_row.cells.len {
		return none
	}
	return target_row.cells[location.col]
}

pub fn (sheet Sheet) get_all_data() !DataFrame {
	return sheet.get_data(sheet.top_left, sheet.bottom_right)
}

pub fn (sheet Sheet) get_data(top_left Location, bottom_right Location) !DataFrame {
	if top_left.row == 0 && bottom_right.row == 0 && sheet.rows.len == 0 {
		return DataFrame{}
	}
	if top_left.row >= sheet.rows.len {
		return error('top_left.row out of range')
	}
	if bottom_right.row > sheet.rows.len {
		return error('bottom_right.row out of range')
	}
	if top_left.col >= sheet.rows[top_left.row].cells.len {
		return error('top_left.col out of range')
	}
	if bottom_right.col > sheet.rows[bottom_right.row].cells.len {
		return error('bottom_right.col out of range')
	}
	mut row_values := [][]string{cap: bottom_right.row - top_left.row + 1}

	for index in top_left.row .. bottom_right.row + 1 {
		row := sheet.rows[index]
		mut cell_values := []string{cap: bottom_right.col - top_left.col + 1}
		for column in top_left.col .. bottom_right.col + 1 {
			cell_values << row.cells[column].value
		}
		row_values << cell_values
	}

	return DataFrame{
		raw_data: row_values
	}
}
