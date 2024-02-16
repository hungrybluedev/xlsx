module main

import xlsx

fn test_empty() ! {
	empty_xml := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>'
	content_types := xlsx.ContentTypes.parse(empty_xml)!
	expected_types := xlsx.ContentTypes{
		defaults: [
			xlsx.DefaultContentType{
				extension: 'rels'
				content_type: 'application/vnd.openxmlformats-package.relationships+xml'
			},
			xlsx.DefaultContentType{
				extension: 'xml'
				content_type: 'application/xml'
			},
		]
		overrides: [
			xlsx.OverrideContentType{
				part_name: '/xl/workbook.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
			},
			xlsx.OverrideContentType{
				part_name: '/xl/worksheets/sheet1.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
			},
			xlsx.OverrideContentType{
				part_name: '/xl/theme/theme1.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.theme+xml'
			},
			xlsx.OverrideContentType{
				part_name: '/xl/styles.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
			},
			xlsx.OverrideContentType{
				part_name: '/docProps/core.xml'
				content_type: 'application/vnd.openxmlformats-package.core-properties+xml'
			},
			xlsx.OverrideContentType{
				part_name: '/docProps/app.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
			},
		]
	}

	assert content_types == expected_types
	assert content_types.str() == empty_xml
}

fn test_sample_data() {
	data_contents := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>'

	content_types := xlsx.ContentTypes.parse(data_contents)!
	expected_types := xlsx.ContentTypes{
		defaults: [
			xlsx.DefaultContentType{
				extension: 'rels'
				content_type: 'application/vnd.openxmlformats-package.relationships+xml'
			},
			xlsx.DefaultContentType{
				extension: 'xml'
				content_type: 'application/xml'
			},
		]
		overrides: [
			xlsx.OverrideContentType{
				part_name: '/xl/workbook.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
			},
			xlsx.OverrideContentType{
				part_name: '/xl/worksheets/sheet1.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
			},
			xlsx.OverrideContentType{
				part_name: '/xl/worksheets/sheet2.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
			},
			xlsx.OverrideContentType{
				part_name: '/xl/theme/theme1.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.theme+xml'
			},
			xlsx.OverrideContentType{
				part_name: '/xl/styles.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
			},
			xlsx.OverrideContentType{
				part_name: '/xl/sharedStrings.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
			},
			xlsx.OverrideContentType{
				part_name: '/docProps/core.xml'
				content_type: 'application/vnd.openxmlformats-package.core-properties+xml'
			},
			xlsx.OverrideContentType{
				part_name: '/docProps/app.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
			},
		]
	}

	assert content_types == expected_types
	assert content_types.str() == data_contents
}
