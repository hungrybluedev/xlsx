import xlsx
import os

fn test_empty() ! {
	path := os.join_path(os.dir(@FILE), 'empty.xlsx')

	document := xlsx.Document.from_file(path)!

	dump(document)
	assert false
}
