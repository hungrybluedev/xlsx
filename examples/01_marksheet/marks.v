import os
import xlsx

fn main() {
	workbook := xlsx.Document.from_file(os.resource_abs_path('data.xlsx'))!
	println('[info] Successfully loaded workbook with ${workbook.sheets.len} worksheets.')

	println('\nAvailable sheets:')
	// sheets are stored as a map, so we can iterate over the keys.
	for index, key in workbook.sheets.keys() {
		println('${index + 1}: "${key}"')
	}

	// Excel uses 1-based indexing for sheets.
	sheet1 := workbook.sheets[1]

	// Note that the Cell struct is able to the CellType.
	// So we can have an idea of what to expect before getting all
	// the data as a dataset with just string data.
	dataset := sheet1.get_all_data()!

	count := dataset.row_count()

	println('\n[info] Sheet 1 has ${count} rows.')

	headers := dataset.raw_data[0]

	println('\nThe headers are:')
	for index, header in headers {
		println('${index + 1}. ${header}')
	}

	println('\nThe student names are:')

	for index in 1 .. count {
		row := dataset.raw_data[index]
		// All data is stored as strings, so we need to convert it to the appropriate type.
		roll := row[0].int()
		name := row[1] + ' ' + row[2]
		println('${roll:02d}. ${name}')
	}
}
