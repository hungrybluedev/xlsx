import xlsx
import os

fn test_table() ! {
	path := os.join_path(os.dir(@FILE), 'table.xlsx')
	document := xlsx.Document.from_file(path)!

	dump(document)
	assert false
}
