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
	created_by  string
	modified_by string
	created_at  time.Time
	modified_at time.Time
}

pub fn (props CoreProperties) str() string {
	time_creation := props.created_at.ymmdd() + 'T' + props.created_at.hhmmss() + 'Z'
	time_modified := props.modified_at.ymmdd() + 'T' + props.modified_at.hhmmss() + 'Z'
	return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"xmlns:dc="http://purl.org/dc/elements/1.1/"xmlns:dcterms="http://purl.org/dc/terms/"xmlns:dcmitype="http://purl.org/dc/dcmitype/"xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>${props.created_by}</dc:creator><cp:lastModifiedBy>${props.modified_by}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${time_creation}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${time_modified}</dcterms:modified></cp:coreProperties>'
}

fn extract_first_element_by_tag(node xml.XMLNode, tag string) !xml.XMLNode {
	tags := node.get_elements_by_tag(tag)
	if tags.len != 1 {
		return error('Invalid core properties XML. Expected a single <${tag}> element.')
	}
	return tags[0]
}

fn extract_first_child_as_string(node xml.XMLNode) !string {
	if node.children.len != 1 {
		return error('Invalid core properties XML. Expected a single child in <${node.name}> element.')
	}
	if node.children[0] !is string {
		return error('Invalid core properties XML. Expected a string child in <${node.name}> element.')
	}
	return node.children[0] as string
}

pub fn CoreProperties.parse(content string) !CoreProperties {
	doc := xml.XMLDocument.from_string(content) or {
		return error('Failed to parse core properties XML.')
	}

	core_properties_nodes := doc.get_elements_by_tag('cp:coreProperties')
	if core_properties_nodes.len != 1 {
		return error('Invalid core properties XML. Expected a single <cp:coreProperties> element.')
	}
	core_properties_node := core_properties_nodes[0]

	mut created_by := ''
	mut modified_by := ''
	mut created_at := time.Time{}
	mut modified_at := time.Time{}

	creator_tags := extract_first_element_by_tag(core_properties_node, 'dc:creator')!
	created_by = extract_first_child_as_string(creator_tags)!

	modified_by_tags := extract_first_element_by_tag(core_properties_node, 'cp:lastModifiedBy')!
	modified_by = extract_first_child_as_string(modified_by_tags)!

	created_at_tags := extract_first_element_by_tag(core_properties_node, 'dcterms:created')!
	created_at = time.parse_iso8601(extract_first_child_as_string(created_at_tags)!) or {
		return error('Invalid core properties XML. Failed to parse created time.')
	}

	modified_at_tags := extract_first_element_by_tag(core_properties_node, 'dcterms:modified')!
	modified_at = time.parse_iso8601(extract_first_child_as_string(modified_at_tags)!) or {
		return error('Invalid core properties XML. Failed to parse modified time.')
	}

	if created_at > modified_at {
		return error('Invalid core properties XML. Created time is newer than modified time.')
	}

	return CoreProperties{created_by, modified_by, created_at, modified_at}
}
