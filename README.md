# xlsx

## Description

A package in pure V for reading and writing Excel files in the XLSX format.

## Roadmap

- [x] Read XLSX files.
- [x] Write XLSX files (basic support).

## Installation

```bash
v install https://github.com/hungrybluedev/xlsx
```

## Usage

### Reading XLSX files

Take the `data.xlsx` file from the `examples/01_marksheet` directory for this example.

```v ignore
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
```

Replace `'data.xlsx'` with the actual path to your file.

For a more detailed example, see [`examples/01_marksheet/marks.v`](examples/01_marksheet/marks.v).

### Writing XLSX files

```v ignore
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
```

For more examples including background fills and styled formulas, see the `spec/` directory.

## Get Involved

- It is a good idea to have examples files ready for testing.
  Ideally, the test files should be as small as possible.
- If it is a feature request, please provide a detailed description
  of the feature and how it should work.

### On GitHub

1. Create issues for bugs you find or features you want to see.
2. Fork the repository and create pull requests for contributions.

### On Discord

1. Join the V Discord server: https://discord.gg/vlang
2. Write in the `#xlsx` channel about your ideas and what you want to do.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for more details.

## Support

If you like this project, please check out [Set Theory for Beginners](https://coderscompass.org/books/set-theory-for-beginners/) or read a few articles on the [Coders' Compass website](https://coderscompass.org/articles).

## Resources

1. [Excel specifications and limits.](https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)
2. [Test Data for sample XLSX files.](https://freetestdata.com/document-files/xlsx/)
