module main

import xlsx

fn test_location_conversion() ! {
	pairs := {
		'A1':         xlsx.Location{
			row:       0
			col:       0
			row_label: '1'
			col_label: 'A'
		}
		'B2':         xlsx.Location{
			row:       1
			col:       1
			row_label: '2'
			col_label: 'B'
		}
		'Z26':        xlsx.Location{
			row:       25
			col:       25
			row_label: '26'
			col_label: 'Z'
		}
		'AA27':       xlsx.Location{
			row:       26
			col:       26
			row_label: '27'
			col_label: 'AA'
		}
		'XFD1048576': xlsx.Location{
			row:       1048575
			col:       16383
			row_label: '1048576'
			col_label: 'XFD'
		}
	}

	for label, location in pairs {
		assert xlsx.Location.from_cartesian(location.row, location.col)! == location
		assert xlsx.Location.from_encoding(label)! == location
	}
}
