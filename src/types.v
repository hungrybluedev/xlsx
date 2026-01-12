module xlsx

pub struct Document {
pub mut:
	shared_strings []string
	sheets         map[int]Sheet
}

pub struct Location {
pub:
	row       int
	col       int
	row_label string
	col_label string
}

pub struct Dimension {
pub:
	top_left     Location
	bottom_right Location
}

pub struct Sheet {
	Dimension
pub:
	name string
pub mut:
	rows []Row
}

pub struct Row {
pub:
	row_index int
	row_label string
pub mut:
	cells []Cell
}

pub enum CellType {
	string_type
	number_type
}

// Currency represents supported currency formats for cell formatting
pub enum Currency {
	usd // US Dollar ($)
	gbp // British Pound (£)
	eur // Euro (€)
	jpy // Japanese Yen (¥)
	cny // Chinese Yuan (¥)
	inr // Indian Rupee (₹)
}

// Returns the Excel format code for this currency
// Format codes use locale IDs for portability across systems
pub fn (c Currency) format_code() string {
	return match c {
		.usd { '[$$-409]#,##0.00' }
		.gbp { '[$£-809]#,##0.00' }
		.eur { '[$€-407]#,##0.00' }
		.jpy { '[$¥-411]#,##0' } // No decimals for Yen
		.cny { '[$¥-804]#,##0.00' }
		.inr { '[$₹-4009]#,##0.00' }
	}
}

// Returns the currency symbol
pub fn (c Currency) symbol() string {
	return match c {
		.usd { '$' }
		.gbp { '£' }
		.eur { '€' }
		.jpy { '¥' }
		.cny { '¥' }
		.inr { '₹' }
	}
}

pub fn CellType.from_code(code string) !CellType {
	match code {
		's' {
			return CellType.string_type
		}
		'n' {
			return CellType.number_type
		}
		else {
			return error('Unknown cell type code: ' + code)
		}
	}
}

pub struct Cell {
pub:
	cell_type CellType
	location  Location
	value     string
	formula   string    // Optional: formula expression (e.g., "C4*D4")
	style_id  int       // Style index for formatting (0=default, 1=date)
	currency  ?Currency // Optional: currency for currency-formatted cells
}

pub struct DataFrame {
pub:
	raw_data [][]string
}

pub fn (data DataFrame) row_count() int {
	return data.raw_data.len
}
