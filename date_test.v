module xlsx

import time

// Test conversion from time.Time to Excel serial date
fn test_time_to_excel_date_jan_1_2026() {
	t := time.Time{
		year:  2026
		month: 1
		day:   1
	}
	result := time_to_excel_date(t)
	assert result == 46023, 'Jan 1, 2026 should be 46023, got ${result}'
}

fn test_time_to_excel_date_jan_8_2026() {
	t := time.Time{
		year:  2026
		month: 1
		day:   8
	}
	result := time_to_excel_date(t)
	assert result == 46030, 'Jan 8, 2026 should be 46030, got ${result}'
}

fn test_time_to_excel_date_jan_15_2026() {
	t := time.Time{
		year:  2026
		month: 1
		day:   15
	}
	result := time_to_excel_date(t)
	assert result == 46037, 'Jan 15, 2026 should be 46037, got ${result}'
}

fn test_time_to_excel_date_jan_22_2026() {
	t := time.Time{
		year:  2026
		month: 1
		day:   22
	}
	result := time_to_excel_date(t)
	assert result == 46044, 'Jan 22, 2026 should be 46044, got ${result}'
}

fn test_time_to_excel_date_jan_29_2026() {
	t := time.Time{
		year:  2026
		month: 1
		day:   29
	}
	result := time_to_excel_date(t)
	assert result == 46051, 'Jan 29, 2026 should be 46051, got ${result}'
}

// Test Excel's 1900 leap year bug boundary
fn test_time_to_excel_date_feb_28_1900() {
	t := time.Time{
		year:  1900
		month: 2
		day:   28
	}
	result := time_to_excel_date(t)
	assert result == 59, 'Feb 28, 1900 should be 59, got ${result}'
}

fn test_time_to_excel_date_march_1_1900() {
	t := time.Time{
		year:  1900
		month: 3
		day:   1
	}
	result := time_to_excel_date(t)
	assert result == 61, 'March 1, 1900 should be 61, got ${result}'
}

fn test_time_to_excel_date_jan_1_1900() {
	t := time.Time{
		year:  1900
		month: 1
		day:   1
	}
	result := time_to_excel_date(t)
	assert result == 1, 'Jan 1, 1900 should be 1, got ${result}'
}

// Test conversion from Excel serial date to time.Time
fn test_excel_date_to_time_jan_1_2026() {
	result := excel_date_to_time(46023)
	assert result.year == 2026, 'Year should be 2026, got ${result.year}'
	assert result.month == 1, 'Month should be 1, got ${result.month}'
	assert result.day == 1, 'Day should be 1, got ${result.day}'
}

fn test_excel_date_to_time_march_1_1900() {
	result := excel_date_to_time(61)
	assert result.year == 1900, 'Year should be 1900, got ${result.year}'
	assert result.month == 3, 'Month should be 3, got ${result.month}'
	assert result.day == 1, 'Day should be 1, got ${result.day}'
}

fn test_excel_date_to_time_feb_28_1900() {
	result := excel_date_to_time(59)
	assert result.year == 1900, 'Year should be 1900, got ${result.year}'
	assert result.month == 2, 'Month should be 2, got ${result.month}'
	assert result.day == 28, 'Day should be 28, got ${result.day}'
}

// Roundtrip tests
fn test_roundtrip_jan_1_2026() {
	original := time.Time{
		year:  2026
		month: 1
		day:   1
	}
	excel_date := time_to_excel_date(original)
	roundtrip := excel_date_to_time(excel_date)
	assert roundtrip.year == original.year, 'Year roundtrip failed'
	assert roundtrip.month == original.month, 'Month roundtrip failed'
	assert roundtrip.day == original.day, 'Day roundtrip failed'
}

fn test_roundtrip_june_15_2024() {
	original := time.Time{
		year:  2024
		month: 6
		day:   15
	}
	excel_date := time_to_excel_date(original)
	roundtrip := excel_date_to_time(excel_date)
	assert roundtrip.year == original.year, 'Year roundtrip failed'
	assert roundtrip.month == original.month, 'Month roundtrip failed'
	assert roundtrip.day == original.day, 'Day roundtrip failed'
}

fn test_roundtrip_dec_31_2099() {
	original := time.Time{
		year:  2099
		month: 12
		day:   31
	}
	excel_date := time_to_excel_date(original)
	roundtrip := excel_date_to_time(excel_date)
	assert roundtrip.year == original.year, 'Year roundtrip failed'
	assert roundtrip.month == original.month, 'Month roundtrip failed'
	assert roundtrip.day == original.day, 'Day roundtrip failed'
}
