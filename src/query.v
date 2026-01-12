module xlsx

pub fn (sheet Sheet) get_cell(location Location) ?Cell {
	// Find row by row_index (not array index) to handle sparse rows
	for row in sheet.rows {
		if row.row_index == location.row {
			// Find cell by column index
			for cell in row.cells {
				if cell.location.col == location.col {
					return cell
				}
			}
			return none // Row exists but cell doesn't
		}
	}
	return none // Row doesn't exist (sparse)
}

pub fn (sheet Sheet) get_all_data() !DataFrame {
	return sheet.get_data(sheet.top_left, sheet.bottom_right)
}

pub fn (sheet Sheet) get_data(top_left Location, bottom_right Location) !DataFrame {
	if top_left.row == 0 && bottom_right.row == 0 && sheet.rows.len == 0 {
		return DataFrame{}
	}

	// Build a map from row_index to array index for efficient lookup
	mut row_index_map := map[int]int{}
	for i, row in sheet.rows {
		row_index_map[row.row_index] = i
	}

	mut row_values := [][]string{cap: bottom_right.row - top_left.row + 1}

	for row_index in top_left.row .. bottom_right.row + 1 {
		mut cell_values := []string{cap: bottom_right.col - top_left.col + 1}

		if row_index in row_index_map {
			row := sheet.rows[row_index_map[row_index]]

			// Build cell lookup for this row
			mut cell_col_map := map[int]string{}
			for cell in row.cells {
				cell_col_map[cell.location.col] = cell.value
			}

			for col_index in top_left.col .. bottom_right.col + 1 {
				if col_index in cell_col_map {
					cell_values << cell_col_map[col_index]
				} else {
					cell_values << '' // Sparse cell
				}
			}
		} else {
			// Sparse row - fill with empty strings
			for _ in top_left.col .. bottom_right.col + 1 {
				cell_values << ''
			}
		}

		row_values << cell_values
	}

	return DataFrame{
		raw_data: row_values
	}
}
