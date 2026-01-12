import xlsx
import os

fn test_empty() ! {
	path := os.join_path(os.dir(@FILE), 'column.xlsx')

	document := xlsx.Document.from_file(path)!

	expected_rows := [
		xlsx.Row{
			row_index: 0
			row_label: '1'
			cells:     [
				xlsx.Cell{
					cell_type: .string_type
					location:  xlsx.Location.from_encoding('A1')!
					value:     'Item 1'
				},
			]
		},
		xlsx.Row{
			row_index: 1
			row_label: '2'
			cells:     [
				xlsx.Cell{
					cell_type: .string_type
					location:  xlsx.Location.from_encoding('A2')!
					value:     'Item 2'
				},
			]
		},
		xlsx.Row{
			row_index: 2
			row_label: '3'
			cells:     [
				xlsx.Cell{
					cell_type: .string_type
					location:  xlsx.Location.from_encoding('A3')!
					value:     'Item 3'
				},
			]
		},
		xlsx.Row{
			row_index: 3
			row_label: '4'
			cells:     [
				xlsx.Cell{
					cell_type: .string_type
					location:  xlsx.Location.from_encoding('A4')!
					value:     'Item 4'
				},
			]
		},
		xlsx.Row{
			row_index: 4
			row_label: '5'
			cells:     [
				xlsx.Cell{
					cell_type: .string_type
					location:  xlsx.Location.from_encoding('A5')!
					value:     'Item 5'
				},
			]
		},
	]
	actual_rows := document.sheets[1].rows
	assert expected_rows == actual_rows, 'Data does not match for ${path}'
}
