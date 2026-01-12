module main

import xlsx
import time

fn test_empty() ! {
	empty_xml := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>'
	content_types := xlsx.ContentTypes.parse(empty_xml)!
	expected_types := xlsx.ContentTypes{
		defaults:  [
			xlsx.DefaultContentType{
				extension:    'rels'
				content_type: 'application/vnd.openxmlformats-package.relationships+xml'
			},
			xlsx.DefaultContentType{
				extension:    'xml'
				content_type: 'application/xml'
			},
		]
		overrides: [
			xlsx.OverrideContentType{
				part_name:    '/xl/workbook.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
			},
			xlsx.OverrideContentType{
				part_name:    '/xl/worksheets/sheet1.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
			},
			xlsx.OverrideContentType{
				part_name:    '/xl/theme/theme1.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.theme+xml'
			},
			xlsx.OverrideContentType{
				part_name:    '/xl/styles.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
			},
			xlsx.OverrideContentType{
				part_name:    '/docProps/core.xml'
				content_type: 'application/vnd.openxmlformats-package.core-properties+xml'
			},
			xlsx.OverrideContentType{
				part_name:    '/docProps/app.xml'
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
		defaults:  [
			xlsx.DefaultContentType{
				extension:    'rels'
				content_type: 'application/vnd.openxmlformats-package.relationships+xml'
			},
			xlsx.DefaultContentType{
				extension:    'xml'
				content_type: 'application/xml'
			},
		]
		overrides: [
			xlsx.OverrideContentType{
				part_name:    '/xl/workbook.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
			},
			xlsx.OverrideContentType{
				part_name:    '/xl/worksheets/sheet1.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
			},
			xlsx.OverrideContentType{
				part_name:    '/xl/worksheets/sheet2.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
			},
			xlsx.OverrideContentType{
				part_name:    '/xl/theme/theme1.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.theme+xml'
			},
			xlsx.OverrideContentType{
				part_name:    '/xl/styles.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
			},
			xlsx.OverrideContentType{
				part_name:    '/xl/sharedStrings.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
			},
			xlsx.OverrideContentType{
				part_name:    '/docProps/core.xml'
				content_type: 'application/vnd.openxmlformats-package.core-properties+xml'
			},
			xlsx.OverrideContentType{
				part_name:    '/docProps/app.xml'
				content_type: 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
			},
		]
	}

	assert content_types == expected_types
	assert content_types.str() == data_contents
}

const core_properties_dataset = [
	xlsx.CoreProperties{
		created_by:  'Subhomoy Haldar'
		modified_by: 'Subhomoy Haldar'
		created_at:  time.parse_iso8601('2024-02-10T10:24:19Z') or {
			panic('Failed to parse time.')
		}
		modified_at: time.parse_iso8601('2024-02-10T10:24:36Z') or {
			panic('Failed to parse time.')
		}
	},
	xlsx.CoreProperties{
		created_by:  'Person A'
		modified_by: 'Person B'
		created_at:  time.parse_iso8601('2024-02-10T10:24:19Z') or {
			panic('Failed to parse time.')
		}
		modified_at: time.parse_iso8601('2024-02-15T12:08:10Z') or {
			panic('Failed to parse time.')
		}
	},
]

fn test_core_properties() {
	for data in core_properties_dataset {
		time_creation := data.created_at.ymmdd() + 'T' + data.created_at.hhmmss() + 'Z'
		time_modified := data.modified_at.ymmdd() + 'T' + data.modified_at.hhmmss() + 'Z'

		core_content := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"xmlns:dc="http://purl.org/dc/elements/1.1/"xmlns:dcterms="http://purl.org/dc/terms/"xmlns:dcmitype="http://purl.org/dc/dcmitype/"xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>${data.created_by}</dc:creator><cp:lastModifiedBy>${data.modified_by}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${time_creation}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${time_modified}</dcterms:modified></cp:coreProperties>'

		core_properties := xlsx.CoreProperties.parse(core_content)!
		assert core_properties == data
		assert core_properties.str() == core_content
	}
}

fn test_app_properties() {
	app_content := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Sample Weather Info</vt:lpstr><vt:lpstr>Sample Altitude Info</vt:lpstr></vt:vector></TitlesOfParts><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>'

	app_properties := xlsx.AppProperties.parse(app_content)!
	expected_properties := xlsx.AppProperties{
		application:        'Microsoft Excel'
		doc_security:       '0'
		scale_crop:         false
		heading_pairs:      [
			xlsx.HeadingPair{
				name:  'Worksheets'
				count: 2
			},
		]
		titles_of_parts:    [
			xlsx.TitlesOfParts{'Sample Weather Info'},
			xlsx.TitlesOfParts{'Sample Altitude Info'},
		]
		company:            ''
		links_up_to_date:   false
		shared_doc:         false
		hyperlinks_changed: false
		app_version:        '16.0300'
	}

	assert app_properties == expected_properties
	assert app_properties.str() == app_content
}
