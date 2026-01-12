import xlsx

fn main() {
	workbook := xlsx.Document.from_file('data.xlsx')!
	println('[info] Loaded ${workbook.sheets.len} worksheets.')

	sheet1 := workbook.sheets[1]
	dataset := sheet1.get_all_data()!

	println('Row count: ${dataset.row_count()}')

	for row in dataset.raw_data {
		println(row.join(', '))
	}
}
