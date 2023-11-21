import xlsx { Location }
import os

fn test_table() ! {
	path := os.join_path(os.dir(@FILE), 'table.xlsx')
	document := xlsx.Document.from_file(path)!

	sheet := document.sheets[1]

	expected_data := xlsx.DataFrame{
		raw_data: [
			['Serial Number', 'X', 'Y'],
			['1', '2', '6'],
			['2', '4', '60'],
			['3', '6', '210'],
			['4', '8', '504'],
			['5', '10', '990'],
			['6', '12', '1716'],
			['7', '14', '2730'],
			['8', '16', '4080'],
			['9', '18', '5814'],
			['10', '20', '7980'],
		]
	}

	full_data := sheet.get_all_data()!

	assert full_data == expected_data

	range_data := sheet.get_data(Location.from_encoding('A1')!, Location.from_encoding('C11')!)!

	assert full_data == range_data
}
