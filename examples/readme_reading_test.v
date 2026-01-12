module main

import os
import xlsx

// Test the reading example from README.md to ensure documentation stays accurate
fn test_readme_reading_example() ! {
	// Load the workbook from the bundled data.xlsx fixture
	workbook := xlsx.Document.from_file(os.join_path(os.dir(@FILE), '01_marksheet', 'data.xlsx'))!

	// Verify worksheet count matches README output
	assert workbook.sheets.len == 1, 'Should have 1 worksheet'

	// Excel uses 1-based indexing for sheets
	sheet1 := workbook.sheets[1]

	// Get all data as a DataFrame
	dataset := sheet1.get_all_data()!
	count := dataset.row_count()

	// Verify row count matches README output (11 rows: 1 header + 10 students)
	assert count == 11, 'Sheet 1 should have 11 rows'

	// Verify headers match README documentation
	headers := dataset.raw_data[0]
	assert headers[0] == 'Roll Number', 'First header should be Roll Number'
	assert headers[1] == 'First Name', 'Second header should be First Name'
	assert headers[2] == 'Last Name', 'Third header should be Last Name'
	assert headers[3] == 'Physics', 'Fourth header should be Physics'
	assert headers[4] == 'Chemistry', 'Fifth header should be Chemistry'
	assert headers[5] == 'Biology', 'Sixth header should be Biology'
	assert headers[6] == 'Mathematics', 'Seventh header should be Mathematics'
	assert headers[7] == 'Total', 'Eighth header should be Total'
	assert headers[8] == 'Percentage', 'Ninth header should be Percentage'

	// Verify first student data matches README output
	row1 := dataset.raw_data[1]
	roll := row1[0].int()
	first_name := row1[1]
	last_name := row1[2]

	assert roll == 1, 'First student roll number should be 1'
	assert first_name == 'Priya', 'First student first name should be Priya'
	assert last_name == 'Patel', 'First student last name should be Patel'

	// Verify we can iterate through all students (as shown in README)
	for index in 1 .. count {
		row := dataset.raw_data[index]
		// All data is stored as strings, so we need to convert to appropriate type
		student_roll := row[0].int()
		student_name := row[1] + ' ' + row[2]

		// Verify roll numbers are sequential 1-10
		assert student_roll == index, 'Roll number should match row index'
		// Verify name is non-empty
		assert student_name.len > 1, 'Student name should not be empty'
	}
}

// Test that sheets can be accessed by key iteration (as shown in README)
fn test_sheet_iteration() ! {
	workbook := xlsx.Document.from_file(os.join_path(os.dir(@FILE), '01_marksheet', 'data.xlsx'))!

	// Verify we can iterate over sheet keys
	keys := workbook.sheets.keys()
	assert keys.len == 1, 'Should have 1 sheet key'

	// Access sheet by the key
	sheet := workbook.sheets[keys[0]]
	assert sheet.name.len > 0, 'Sheet should have a name'
}
