import xlsx
import time

fn main() {
	mut doc := xlsx.Document.new()

	sheet_id := doc.add_sheet('Sheet1')
	mut sheet := doc.get_sheet_mut(sheet_id) or { return }

	sheet.set_cell(xlsx.Location.from_encoding('A1')!, 'Name')
	sheet.set_cell(xlsx.Location.from_encoding('B1')!, 'Score')

	sheet.set_number(xlsx.Location.from_encoding('B2')!, 95)
	sheet.set_number_f64(xlsx.Location.from_encoding('B3')!, 87.5)

	jan_1_2026 := time.Time{
		year:  2026
		month: 1
		day:   1
	}
	sheet.set_date(xlsx.Location.from_encoding('C2')!, jan_1_2026)

	sheet.set_formula(xlsx.Location.from_encoding('B4')!, 'SUM(B2:B3)')
	sheet.set_currency(xlsx.Location.from_encoding('D2')!, 1234.56, .usd)

	doc.to_file('output.xlsx')!
}
