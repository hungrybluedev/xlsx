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

	creator_tag := extract_first_element_by_tag(core_properties_node, 'dc:creator')!
	created_by = extract_first_child_as_string(creator_tag)!

	modified_by_tag := extract_first_element_by_tag(core_properties_node, 'cp:lastModifiedBy')!
	modified_by = extract_first_child_as_string(modified_by_tag)!

	created_at_tag := extract_first_element_by_tag(core_properties_node, 'dcterms:created')!
	created_at = time.parse_iso8601(extract_first_child_as_string(created_at_tag)!) or {
		return error('Invalid core properties XML. Failed to parse created time.')
	}

	modified_at_tag := extract_first_element_by_tag(core_properties_node, 'dcterms:modified')!
	modified_at = time.parse_iso8601(extract_first_child_as_string(modified_at_tag)!) or {
		return error('Invalid core properties XML. Failed to parse modified time.')
	}

	if created_at > modified_at {
		return error('Invalid core properties XML. Created time is newer than modified time.')
	}

	return CoreProperties{created_by, modified_by, created_at, modified_at}
}

pub struct HeadingPair {
	name  string
	count int
}

pub fn (pair HeadingPair) str() string {
	return '<vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>${pair.name}</vt:lpstr></vt:variant><vt:variant><vt:i4>${pair.count}</vt:i4></vt:variant></vt:vector>'
}

fn HeadingPair.parse(node xml.XMLNode) ![]HeadingPair {
	mut pairs := []HeadingPair{}

	vector_tags := node.get_elements_by_tag('vt:vector')
	for vector_tag in vector_tags {
		variant_tags := vector_tag.get_elements_by_tag('vt:variant')
		if variant_tags.len != 2 {
			return error('Invalid app properties XML. Expected two <vt:variant> elements.')
		}

		name_tag := variant_tags[0].get_elements_by_tag('vt:lpstr')
		if name_tag.len != 1 {
			return error('Invalid app properties XML. Expected a single <vt:lpstr> element.')
		}
		name := name_tag[0].children[0] as string

		count_tag := variant_tags[1].get_elements_by_tag('vt:i4')
		if count_tag.len != 1 {
			return error('Invalid app properties XML. Expected a single <vt:i4> element.')
		}
		count_text := count_tag[0].children[0] as string
		count := count_text.int()

		pairs << HeadingPair{name, count}
	}

	return pairs
}

fn encode_heading_pairs(pairs []HeadingPair) string {
	mut result := strings.new_builder(256)

	result.write_string('<HeadingPairs>')
	for pair in pairs {
		result.write_string(pair.str())
	}
	result.write_string('</HeadingPairs>')

	return result.str()
}

pub struct TitlesOfParts {
	entity string
}

pub fn (title TitlesOfParts) str() string {
	return '<vt:lpstr>${title.entity}</vt:lpstr>'
}

fn encode_titles_of_parts(titles []TitlesOfParts) string {
	mut result := strings.new_builder(128)

	result.write_string('<TitlesOfParts><vt:vector size="${titles.len}" baseType="lpstr">')
	for title in titles {
		result.write_string(title.str())
	}
	result.write_string('</vt:vector></TitlesOfParts>')

	return result.str()
}

fn TitlesOfParts.parse(node xml.XMLNode) ![]TitlesOfParts {
	mut titles := []TitlesOfParts{}

	lpstr_tags := node.get_elements_by_tag('vt:lpstr')
	for tag in lpstr_tags {
		titles << TitlesOfParts{tag.children[0] as string}
	}

	return titles
}

pub struct AppProperties {
	application        string
	doc_security       string
	scale_crop         bool
	links_up_to_date   bool
	shared_doc         bool
	hyperlinks_changed bool
	app_version        string
	company            string
	heading_pairs      []HeadingPair
	titles_of_parts    []TitlesOfParts
}

pub fn (prop AppProperties) str() string {
	return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>${prop.application}</Application><DocSecurity>${prop.doc_security}</DocSecurity><ScaleCrop>${prop.scale_crop}</ScaleCrop>${encode_heading_pairs(prop.heading_pairs)}${encode_titles_of_parts(prop.titles_of_parts)}<Company>${prop.company}</Company><LinksUpToDate>${prop.links_up_to_date}</LinksUpToDate><SharedDoc>${prop.shared_doc}</SharedDoc><HyperlinksChanged>${prop.hyperlinks_changed}</HyperlinksChanged><AppVersion>${prop.app_version}</AppVersion></Properties>'
}

pub fn AppProperties.parse(content string) !AppProperties {
	doc := xml.XMLDocument.from_string(content) or {
		return error('Failed to parse app properties XML.')
	}

	properties_nodes := doc.get_elements_by_tag('Properties')
	if properties_nodes.len != 1 {
		return error('Invalid app properties XML. Expected a single <Properties> element.')
	}

	properties_node := properties_nodes[0]

	application_tag := extract_first_element_by_tag(properties_node, 'Application')!
	application := extract_first_child_as_string(application_tag)!

	doc_security_tag := extract_first_element_by_tag(properties_node, 'DocSecurity')!
	doc_security := extract_first_child_as_string(doc_security_tag)!
	if doc_security != '0' && doc_security != '1' {
		return error('Invalid app properties XML. Expected a "0" or "1" value for <DocSecurity>.')
	}

	scale_crop_tag := extract_first_element_by_tag(properties_node, 'ScaleCrop')!
	scale_crop_text := extract_first_child_as_string(scale_crop_tag)!
	scale_crop := scale_crop_text == 'true'

	links_up_to_date_tag := extract_first_element_by_tag(properties_node, 'LinksUpToDate')!
	links_up_to_date_text := extract_first_child_as_string(links_up_to_date_tag)!
	links_up_to_date := links_up_to_date_text == 'true'

	shared_doc_tag := extract_first_element_by_tag(properties_node, 'SharedDoc')!
	shared_doc_text := extract_first_child_as_string(shared_doc_tag)!
	shared_doc := shared_doc_text == 'true'

	hyperlinks_changed_tag := extract_first_element_by_tag(properties_node, 'HyperlinksChanged')!
	hyperlinks_changed_text := extract_first_child_as_string(hyperlinks_changed_tag)!
	hyperlinks_changed := hyperlinks_changed_text == 'true'

	app_version_tag := extract_first_element_by_tag(properties_node, 'AppVersion')!
	app_version := extract_first_child_as_string(app_version_tag)!

	company_tag := extract_first_element_by_tag(properties_node, 'Company')!
	company := extract_first_child_as_string(company_tag) or { '' }

	heading_pairs_tag := extract_first_element_by_tag(properties_node, 'HeadingPairs')!
	heading_pairs := HeadingPair.parse(heading_pairs_tag) or {
		return error('Invalid app properties XML. Failed to parse heading pairs.\n${err}')
	}

	titles_of_parts_tag := extract_first_element_by_tag(properties_node, 'TitlesOfParts')!
	titles_of_parts := TitlesOfParts.parse(titles_of_parts_tag) or {
		return error('Invalid app properties XML. Failed to parse titles of parts.\n${err}')
	}

	return AppProperties{
		application: application
		doc_security: doc_security
		scale_crop: scale_crop
		links_up_to_date: links_up_to_date
		shared_doc: shared_doc
		hyperlinks_changed: hyperlinks_changed
		app_version: app_version
		company: company
		heading_pairs: heading_pairs
		titles_of_parts: titles_of_parts
	}
}
