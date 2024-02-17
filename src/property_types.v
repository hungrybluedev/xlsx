module xlsx

import encoding.xml
import strings
import time

pub struct DefaultContentType {
pub:
	extension    string
	content_type string
}

pub fn (default DefaultContentType) str() string {
	return '<Default Extension="${default.extension}" ContentType="${default.content_type}"/>'
}

pub struct OverrideContentType {
pub:
	part_name    string
	content_type string
}

pub fn (override OverrideContentType) str() string {
	return '<Override PartName="${override.part_name}" ContentType="${override.content_type}"/>'
}

pub struct ContentTypes {
pub:
	defaults  []DefaultContentType
	overrides []OverrideContentType
}

pub fn (content_type ContentTypes) str() string {
	mut result := strings.new_builder(128)

	result.write_string('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
	result.write_string('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')

	for default in content_type.defaults {
		result.write_string(default.str())
	}
	for override in content_type.overrides {
		result.write_string(override.str())
	}

	result.write_string('</Types>')

	return result.str()
}

pub fn ContentTypes.parse(content string) !ContentTypes {
	mut defaults := []DefaultContentType{}
	mut overrides := []OverrideContentType{}

	doc := xml.XMLDocument.from_string(content) or {
		return error('Failed to parse content types XML.')
	}

	content_types_node := doc.get_elements_by_tag('Types')
	if content_types_node.len != 1 {
		return error('Invalid content types XML. Expected a single <Types> element.')
	}

	default_tags := content_types_node[0].get_elements_by_tag('Default')
	if default_tags.len < 2 {
		return error('Invalid content types XML. Expected at least two <Default> elements.')
	}
	for tag in default_tags {
		if 'Extension' !in tag.attributes {
			return error("Invalid content types XML. Expected an 'Extension' attribute in <Default> element.")
		}
		if 'ContentType' !in tag.attributes {
			return error("Invalid content types XML. Expected a 'ContentType' attribute in <Default> element.")
		}
		defaults << DefaultContentType{tag.attributes['Extension'], tag.attributes['ContentType']}
	}

	override_tags := content_types_node[0].get_elements_by_tag('Override')
	for tag in override_tags {
		if 'PartName' !in tag.attributes {
			return error("Invalid content types XML. Expected a 'PartName' attribute in <Override> element.")
		}
		if 'ContentType' !in tag.attributes {
			return error("Invalid content types XML. Expected a 'ContentType' attribute in <Override> element.")
		}
		overrides << OverrideContentType{tag.attributes['PartName'], tag.attributes['ContentType']}
	}

	return ContentTypes{defaults, overrides}
}

pub struct CoreProperties {
pub:
	created_by string
	modified_by string
	created_at time.Time
	modified_at time.Time
}

pub fn (props CoreProperties) str() string {
	time_creation := props.created_at.ymmdd() + 'T' + props.created_at.hhmmss() + 'Z'
	return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"xmlns:dc="http://purl.org/dc/elements/1.1/"xmlns:dcterms="http://purl.org/dc/terms/"xmlns:dcmitype="http://purl.org/dc/dcmitype/"xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>${props.created_by}</dc:creator><cp:lastModifiedBy>${props.modified_by}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2024-02-10T10:24:19Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2024-02-10T10:24:36Z</dcterms:modified></cp:coreProperties>'
}
