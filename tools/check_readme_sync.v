module main

import os

fn main() {
	root := os.dir(os.dir(@FILE))

	readme := os.read_file(os.join_path(root, 'README.md')) or {
		eprintln('ERROR: Failed to read README.md: ${err}')
		exit(1)
	}

	snippets := {
		'reading': os.join_path(root, 'examples', 'snippets', 'reading_example.v')
		'writing': os.join_path(root, 'examples', 'snippets', 'writing_example.v')
	}

	mut passed := true
	for name, path in snippets {
		expected := os.read_file(path) or {
			eprintln('ERROR: Failed to read ${path}: ${err}')
			exit(1)
		}

		if !readme.contains(expected.trim_space()) {
			eprintln('FAIL: ${name} example not found in README (exact match required)')
			eprintln('      Source: ${path}')
			passed = false
		} else {
			println('PASS: ${name} example matches')
		}
	}

	if passed {
		println('\nAll README examples are in sync.')
	} else {
		eprintln('\nREADME examples are OUT OF SYNC with source files.')
		eprintln('Update README.md to match the snippet files exactly.')
		exit(1)
	}
}
