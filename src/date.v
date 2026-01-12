module xlsx

import time

// Excel date serial number for January 1, 1970 (Unix epoch)
// Excel uses a serial date system where day 1 = January 1, 1900
// Due to a historical bug, Excel incorrectly treats 1900 as a leap year
// (Feb 29, 1900 = serial 60, which doesn't exist in reality)
const unix_epoch_as_excel_serial = 25569

// Excel serial number for the fake February 29, 1900
const excel_fake_leap_day = 60

// Converts a time.Time to an Excel serial date number.
// Handles Excel's 1900 leap year bug correctly.
pub fn time_to_excel_date(t time.Time) int {
	days_since_unix := time.days_from_unix_epoch(t.year, t.month, t.day)
	excel_date := days_since_unix + unix_epoch_as_excel_serial

	// For dates before March 1, 1900 (serial < 61), subtract 1
	// because these dates don't include the fake Feb 29, 1900
	if excel_date < 61 {
		return excel_date - 1
	}
	return excel_date
}

// Converts an Excel serial date number to a time.Time.
// Handles Excel's 1900 leap year bug correctly.
// Note: Serial 60 (Excel's fake Feb 29, 1900) maps to March 1, 1900.
pub fn excel_date_to_time(excel_date int) time.Time {
	mut days_since_unix := 0

	// Account for Excel's fake Feb 29, 1900
	// For dates after the fake day (serial > 60), subtract 25569
	// For dates at or before (serial <= 60), subtract 25568
	if excel_date > excel_fake_leap_day {
		days_since_unix = excel_date - unix_epoch_as_excel_serial
	} else {
		days_since_unix = excel_date - unix_epoch_as_excel_serial + 1
	}

	return time.date_from_days_after_unix_epoch(days_since_unix)
}
