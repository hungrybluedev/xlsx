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

```v
import xlsx

fn main() {
	workbook := xlsx.Document.from_file('path/to/data.xlsx')!
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
```

Remember to replace `'path/to/data.xlsx'` with the actual path to the file.

After you are done, run the program:

```bash
v run marksheet.v
```

You should see the following output:

```plaintext
[info] Successfully loaded workbook with 1 worksheets.

Available sheets:
1: "1"

[info] Sheet 1 has 11 rows.

The headers are:
1. Roll Number
2. First Name
3. Last Name
4. Physics
5. Chemistry
6. Biology
7. Mathematics
8. Total
9. Percentage

The student names are:
01. Priya Patel
02. Kwame Nkosi
03. Mei Chen
04. Aisha Adekunle
05. Javed Khan
06. Mei-Ling Wong
07. Oluwafemi Adeyemi
08. Yuki Takahashi
09. Rashid Al-Mansoori
10. Sanya Verma
```

Try running the example on other XLSX files to see how it works.
Modify the example to suit your needs.

### Writing XLSX files

```v
import xlsx

fn main() {
    // Create a new document
    mut doc := xlsx.Document.new()

    // Add a sheet and get a mutable reference
    sheet_id := doc.add_sheet('Sheet1')
    mut sheet := doc.get_sheet_mut(sheet_id) or { return }

    // Set string cells
    sheet.set_cell(xlsx.Location.from_encoding('A1')!, 'Name')
    sheet.set_cell(xlsx.Location.from_encoding('B1')!, 'Score')

    // Set numeric cells
    sheet.set_number(xlsx.Location.from_encoding('B2')!, 95)
    sheet.set_number_f64(xlsx.Location.from_encoding('B3')!, 87.5)

    // Set dates (Excel serial date format)
    sheet.set_date(xlsx.Location.from_encoding('C2')!, 46023) // Jan 1, 2026

    // Set formulas
    sheet.set_formula(xlsx.Location.from_encoding('B4')!, 'SUM(B2:B3)')

    // Set currency values (supports usd, gbp, eur, jpy, cny, inr)
    sheet.set_currency(xlsx.Location.from_encoding('D2')!, 1234.56, .usd)

    // Save to file
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

If you like this project, please consider supporting me on [GitHub Sponsors](https://github.com/sponsors/hungrybluedev).

## Resources

1. [Excel specifications and limits.](https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)
2. [Test Data for sample XLSX files.](https://freetestdata.com/document-files/xlsx/)
