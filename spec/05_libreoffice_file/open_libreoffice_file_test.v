import os
import xlsx

fn test_opening_a_libreoffice_calc_table() {
	workbook := xlsx.Document.from_file(os.join_path(os.dir(@FILE), 'abc.xlsx'))!
	println('[info] Successfully loaded workbook with ${workbook.sheets.len} worksheets.')
	println('\nAvailable sheets:')
	for index, key in workbook.sheets.keys() {
		println('   sheet ${index + 1} has key: "${key}"')
	}
	sheet1 := workbook.sheets[1]
	dataset := sheet1.get_all_data()!
	count := dataset.row_count()
	println('\n[info] Sheet 1 has ${count} rows.')
	headers := dataset.raw_data[0]
	println('\nThe headers are:')
	assert headers.len == 1
	for index, header in headers {
		println('${index + 1}. ${header}')
		assert index == 0
		assert header == 'abc'
	}
}
