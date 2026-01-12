module xlsx

// Test that get_data handles sparse rows (rows with gaps in row_index)
fn test_get_data_with_sparse_rows() ! {
	// Create a sheet with rows 0, 1, and 3 (row 2 is missing - sparse)
	sheet := Sheet{
		name:         'TestSheet'
		rows:         [
			Row{
				row_index: 0
				row_label: '1'
				cells:     [
					Cell{
						cell_type: .string_type
						location:  Location{
							row:       0
							col:       0
							row_label: '1'
							col_label: 'A'
						}
						value:     'A1'
					},
					Cell{
						cell_type: .string_type
						location:  Location{
							row:       0
							col:       1
							row_label: '1'
							col_label: 'B'
						}
						value:     'B1'
					},
				]
			},
			Row{
				row_index: 1
				row_label: '2'
				cells:     [
					Cell{
						cell_type: .string_type
						location:  Location{
							row:       1
							col:       0
							row_label: '2'
							col_label: 'A'
						}
						value:     'A2'
					},
					Cell{
						cell_type: .string_type
						location:  Location{
							row:       1
							col:       1
							row_label: '2'
							col_label: 'B'
						}
						value:     'B2'
					},
				]
			},
			// Row 2 is missing (sparse)
			Row{
				row_index: 3
				row_label: '4'
				cells:     [
					Cell{
						cell_type: .string_type
						location:  Location{
							row:       3
							col:       0
							row_label: '4'
							col_label: 'A'
						}
						value:     'A4'
					},
					Cell{
						cell_type: .string_type
						location:  Location{
							row:       3
							col:       1
							row_label: '4'
							col_label: 'B'
						}
						value:     'B4'
					},
				]
			},
		]
		top_left:     Location{
			row:       0
			col:       0
			row_label: '1'
			col_label: 'A'
		}
		bottom_right: Location{
			row:       3
			col:       1
			row_label: '4'
			col_label: 'B'
		}
	}

	// This should handle the sparse row gracefully
	data := sheet.get_all_data()!

	// Should have 4 rows (0, 1, 2, 3)
	assert data.raw_data.len == 4, 'should have 4 rows, got ${data.raw_data.len}'

	// Row 0
	assert data.raw_data[0][0] == 'A1'
	assert data.raw_data[0][1] == 'B1'

	// Row 1
	assert data.raw_data[1][0] == 'A2'
	assert data.raw_data[1][1] == 'B2'

	// Row 2 (sparse - should be empty strings)
	assert data.raw_data[2][0] == '', 'sparse row should have empty string'
	assert data.raw_data[2][1] == '', 'sparse row should have empty string'

	// Row 3
	assert data.raw_data[3][0] == 'A4'
	assert data.raw_data[3][1] == 'B4'
}

// Test that get_cell handles sparse rows
fn test_get_cell_with_sparse_rows() ! {
	sheet := Sheet{
		name:         'TestSheet'
		rows:         [
			Row{
				row_index: 0
				row_label: '1'
				cells:     [
					Cell{
						cell_type: .string_type
						location:  Location{
							row:       0
							col:       0
							row_label: '1'
							col_label: 'A'
						}
						value:     'A1'
					},
				]
			},
			// Row 1 is missing
			Row{
				row_index: 2
				row_label: '3'
				cells:     [
					Cell{
						cell_type: .string_type
						location:  Location{
							row:       2
							col:       0
							row_label: '3'
							col_label: 'A'
						}
						value:     'A3'
					},
				]
			},
		]
		top_left:     Location{
			row:       0
			col:       0
			row_label: '1'
			col_label: 'A'
		}
		bottom_right: Location{
			row:       2
			col:       0
			row_label: '3'
			col_label: 'A'
		}
	}

	// Existing cell should be found
	cell_a1 := sheet.get_cell(Location{ row: 0, col: 0, row_label: '1', col_label: 'A' }) or {
		return error('A1 should exist')
	}
	assert cell_a1.value == 'A1'

	// Sparse row cell should return none
	cell_a2 := sheet.get_cell(Location{ row: 1, col: 0, row_label: '2', col_label: 'A' })
	assert cell_a2 == none, 'sparse row cell should return none'

	// Cell in row after gap should be found
	cell_a3 := sheet.get_cell(Location{ row: 2, col: 0, row_label: '3', col_label: 'A' }) or {
		return error('A3 should exist')
	}
	assert cell_a3.value == 'A3'
}
