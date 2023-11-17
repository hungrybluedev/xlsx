module xlsx

pub struct Document {
pub:
	shared_strings []string
	sheets         map[int]Sheet
}

pub struct Sheet {
pub:
	name string
	rows []Row
}

pub struct Row {
pub:
	row_index int
	row_label string
	cells     []Cell
}

pub enum CellType {
	string_type
	number_type
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

pub struct Location {
pub:
	row       int
	col       int
	row_label string
	col_label string
}

pub struct Cell {
pub:
	cell_type CellType
	location  Location
	value     string
}

pub struct DataFrame {
	
}
