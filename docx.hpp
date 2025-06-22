/*
This file is part of Simple Office Open XML Document (docx) Library.

Simple Office Open XML Document (docx) Library is free software: you can redistribute it and/or modify it under
the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of
the License, or (at your option) any later version.

Simple Office Open XML Document (docx) Library is distributed in the hope that it will be useful, but WITHOUT
ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with Simple Office Open XML Document (docx) Library.
If not, see <https://www.gnu.org/licenses/>.
*/

#ifndef DOCX_HPP
#define DOCX_HPP

#include "../xml/xml.hpp"

#include <string>
#include <cstdlib>
#include <filesystem>

inline constexpr const char* newl = "\n";

//////////////////////
// DOCX declaration //
//////////////////////

class DOCX {
public:
    DOCX() = default;

    class Paragraph;
    class Text;

    void add_paragraph(DOCX::Paragraph paragraph);
    void add_empty_line(size_t count = 1);
    void print();
    void save(std::string fname);
    void set_global_font_size(size_t set_size); // TODO not used yet
    size_t get_global_font_size();

private:
    std::vector<DOCX::Paragraph> paragraphs;

    XML::Node get();
    void generate_docx_template();
    void zip_docx_template(std::string fname);
    void delete_docx_template();

    static size_t global_font_size;
    static XML::Node root_node();
};

///////////////////////////
// Paragraph declaration //
///////////////////////////

class DOCX::Paragraph {
public:
    Paragraph() = default;

    bool empty = false; // if true all content is ignored and a new empty line is shown

    void add_formatted_text(Text t);
    void add_plain_text(std::string text_str);
    void add_space(size_t count = 1);
    void add_bold_text(std::string text_str);
    void add_italic_text(std::string text_str);
    void add_underlined_text(std::string text_str);
    void add_struckthrough_text(std::string text_str);
    XML::Node get();
    
    static XML::Node empty_line_node();

private:
    std::vector<DOCX::Text> contents;
};

//////////////////////
// Text declaration //
//////////////////////

class DOCX::Text {
public:
    Text() = default;
    Text(std::string set_text);
    std::string text;
    bool bold = false;
    bool italic = false;
    bool underline = false;
    bool strikethrough = false;
    bool preserve_space = false;
    size_t size = 12;
};

////////////////////////////
// DOCX Utils declaration //
////////////////////////////

class DOCXUtils {
public:
    static void mkdir(std::string dirpath);
    static void write_file(std::string fpath, std::string content);
    static void delete_file_or_folder(std::string dirpath);

    static std::string content_types_file();
    static std::string dotrels_file();
    static std::string app_file();
    static std::string core_file();
    static std::string font_table_file();
    static std::string settings_file();
    static std::string styles_file();
    static std::string document_xml_rels_file();
    static std::string theme1_file();
};

//////////////////////
// DOCX definitions //
//////////////////////

inline void DOCX::add_paragraph(DOCX::Paragraph paragraph) {
    paragraphs.push_back(paragraph);
}

inline void DOCX::add_empty_line(size_t count) {
    for (size_t i = 0; i < count; i++) {
        DOCX::Paragraph p;
        p.empty = true;
        paragraphs.push_back(p);
    }
}

inline void DOCX::print() {
    get().print();
}

inline void DOCX::save(std::string fname) {
    generate_docx_template();
    XML::Node root = get();
    root.save("docx_temp/word/document.xml");
    zip_docx_template(fname);
    delete_docx_template();
}

// TODO not used yet
inline void DOCX::set_global_font_size(size_t set_size) {
    global_font_size = set_size;
}

inline size_t DOCX::get_global_font_size() {
    return global_font_size;
}

inline XML::Node DOCX::get() {
    XML::Node root = root_node();
    XML::Node body("w:body");

    for (size_t i = 0; i < paragraphs.size(); i++) {
        DOCX::Paragraph cur_p = paragraphs.at(i);
        if (cur_p.empty) {
            body.add_child(DOCX::Paragraph::empty_line_node());
        } else {
            body.add_child(paragraphs.at(i).get());
        }
    }

    root.add_child(body);
    return root;
}

inline size_t DOCX::global_font_size = 12;

inline XML::Node DOCX::root_node() {
    XML::Node root("w:document");
    root.attributes["xmlns:o"] = "urn:schemas-microsoft-com:office:office";
    root.attributes["xmlns:r"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
	root.attributes["xmlns:v"] = "urn:schemas-microsoft-com:vml";
	root.attributes["xmlns:w"] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
	root.attributes["xmlns:w10"] = "urn:schemas-microsoft-com:office:word";
	root.attributes["xmlns:wp"] = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
	root.attributes["xmlns:pic"] = "http://schemas.openxmlformats.org/drawingml/2006/picture";
	root.attributes["xmlns:wps"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
	root.attributes["xmlns:wpg"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
	root.attributes["xmlns:mc"] = "http://schemas.openxmlformats.org/markup-compatibility/2006";
	root.attributes["xmlns:wp14"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
	root.attributes["xmlns:w14"] = "http://schemas.microsoft.com/office/word/2010/wordml";
	root.attributes["xmlns:w15"] = "http://schemas.microsoft.com/office/word/2012/wordml";
    root.attributes["mc:Ignorable"] = "w14 wp14 w15";
    return root;
}

///////////////////////////
// Paragraph definitions //
///////////////////////////

inline void DOCX::Paragraph::add_formatted_text(Text t) {
    contents.push_back(t);
}

inline void DOCX::Paragraph::add_plain_text(std::string text_str) {
    Text t(text_str);
    contents.push_back(t);
}

inline void DOCX::Paragraph::add_space(size_t count) {
    std::string spaces;
    for (size_t i = 0; i < count; i++) {
        spaces += " ";
    }
    Text t(spaces);
    t.preserve_space = true;
    contents.push_back(t);
}

inline void DOCX::Paragraph::add_bold_text(std::string text_str) {
    Text t(text_str);
    t.bold = true;
    contents.push_back(t);
}

inline void DOCX::Paragraph::add_italic_text(std::string text_str) {
    Text t(text_str);
    t.italic = true;
    contents.push_back(t);
}

inline void DOCX::Paragraph::add_underlined_text(std::string text_str) {
    Text t(text_str);
    t.underline = true;
    contents.push_back(t);
}

inline void DOCX::Paragraph::add_struckthrough_text(std::string text_str) {
    Text t(text_str);
    t.strikethrough = true;
    contents.push_back(t);
}

inline XML::Node DOCX::Paragraph::get() {
    XML::Node p("w:p");
    {
        XML::Node pPr("w:pPr");
        {
            XML::Node pStyle("w:pStyle");
            pStyle.self_closing = true;
            pStyle.attributes["w:val"] = "Normal";
            XML::Node bidi("w:bidi");
            bidi.self_closing = true;
            bidi.attributes["w:val"] = "0";
            XML::Node jc("w:jc");
            jc.self_closing = true;
            jc.attributes["w:val"] = "start";
            XML::Node rPr("w:rPr");

            pPr.add_child(pStyle);
            pPr.add_child(bidi);
            pPr.add_child(jc);
            pPr.add_child(rPr);
        }
        p.add_child(pPr);

        for (size_t i = 0; i < contents.size(); i++) {
            Text cur_text = contents.at(i);

            XML::Node r("w:r");
            {
                XML::Node rPr("w:rPr");
                {
                    if (cur_text.size != DOCX::global_font_size) {
                        XML::Node sz("w:sz");
                        sz.self_closing = true;
                        sz.attributes["w:val"] = std::to_string(cur_text.size * 2); // because half points
                        rPr.add_child(sz);
                    }

                    if (cur_text.bold) {
                        XML::Node b("w:b");
                        b.self_closing = true;
                        XML::Node bCs("w:bCs");
                        bCs.self_closing = true;

                        rPr.add_child(b);
                        rPr.add_child(bCs);
                    }

                    if (cur_text.italic) {
                        XML::Node i("w:i");
                        i.self_closing = true;
                        XML::Node iCs("w:iCs");
                        iCs.self_closing = true;

                        rPr.add_child(i);
                        rPr.add_child(iCs);
                    }

                    if (cur_text.underline) {
                        XML::Node u("w:u");
                        u.attributes["w:val"] = "single";
                        rPr.add_child(u);
                    }

                    if (cur_text.strikethrough) {
                        XML::Node strike("w:strike");
                        strike.self_closing = true;
                        rPr.add_child(strike);
                    }
                }
                r.add_child(rPr);

                XML::Node t("w:t");
                t.content = cur_text.text;
                if (cur_text.preserve_space) {
                    t.attributes["xml:space"] = "preserve";
                }
                r.add_child(t);
            }
            p.add_child(r);
        }
    }
    return p;
}

inline void DOCX::generate_docx_template() {
    const char* PATH_DELIM = "/"; // for linux TODO make cross platform

    // Create root folder
    std::string root_dirpath = "./docx_temp";
    DOCXUtils::mkdir(root_dirpath);

    // Create directories in the root folder
    DOCXUtils::mkdir(root_dirpath + PATH_DELIM + "_rels");
    DOCXUtils::mkdir(root_dirpath + PATH_DELIM + "docProps");
    DOCXUtils::mkdir(root_dirpath + PATH_DELIM + "word");

    // Create directories in the root/word folder
    DOCXUtils::mkdir(root_dirpath + PATH_DELIM + "word" + PATH_DELIM + "_rels");
    DOCXUtils::mkdir(root_dirpath + PATH_DELIM + "word" + PATH_DELIM + "theme");

    // Write root/[Content_Types].xml file
    DOCXUtils::write_file(root_dirpath + PATH_DELIM + "[Content_Types].xml", DOCXUtils::content_types_file());

    // Write root/_rels/.rels file
    DOCXUtils::write_file(root_dirpath + PATH_DELIM + "_rels" + PATH_DELIM + ".rels", DOCXUtils::dotrels_file());

    // Write app.xml and core.xml in root/docProps
    DOCXUtils::write_file(root_dirpath + PATH_DELIM + "docProps" + PATH_DELIM + "app.xml", DOCXUtils::app_file());
    DOCXUtils::write_file(root_dirpath + PATH_DELIM + "docProps" + PATH_DELIM + "core.xml", DOCXUtils::core_file());

    // Write fontTable.xml, settings.xml, styles.xml files in root/word
    DOCXUtils::write_file(root_dirpath + PATH_DELIM + "word" + PATH_DELIM + "fontTable.xml", DOCXUtils::font_table_file());
    DOCXUtils::write_file(root_dirpath + PATH_DELIM + "word" + PATH_DELIM + "settings.xml", DOCXUtils::settings_file());
    DOCXUtils::write_file(root_dirpath + PATH_DELIM + "word" + PATH_DELIM + "styles.xml", DOCXUtils::styles_file());

    // Write root/word/_rels/document.xml.rels
    DOCXUtils::write_file(root_dirpath + PATH_DELIM + "word" + PATH_DELIM + "_rels" + PATH_DELIM + "document.xml.rels", DOCXUtils::document_xml_rels_file());

    // Write root/word/theme/theme1.xml
    DOCXUtils::write_file(root_dirpath + PATH_DELIM + "word" + PATH_DELIM + "theme" + PATH_DELIM + "theme1.xml", DOCXUtils::theme1_file());
}

inline void DOCX::zip_docx_template(std::string fname) {
    std::string save_cmd = "( cd docx_temp && zip -r ../" + fname + " . > /dev/null 2>&1 )";
    system(save_cmd.c_str());
}

inline void DOCX::delete_docx_template() {
    std::vector<std::string> files = {
        "docx_temp/_rels/.rels",
        "docx_temp/_rels",
        "docx_temp/docProps/app.xml",
        "docx_temp/docProps/core.xml",
        "docx_temp/docProps",
        "docx_temp/word/document.xml",
        "docx_temp/word/fontTable.xml",
        "docx_temp/word/settings.xml",
        "docx_temp/word/styles.xml",
        "docx_temp/word/_rels/document.xml.rels",
        "docx_temp/word/theme/theme1.xml",
        "docx_temp/word/_rels",
        "docx_temp/word/theme",
        "docx_temp/word",
        "docx_temp/[Content_Types].xml",
        "docx_temp"
    };

    for (size_t i = 0; i < files.size(); i++) {
        std::string cur_file = files.at(i);
        DOCXUtils::delete_file_or_folder(cur_file);
    }
}

inline XML::Node DOCX::Paragraph::empty_line_node() {
    XML::Node p("w:p");
    {
        XML::Node pPr("w:pPr");
        {
            XML::Node pStyle("w:pStyle");
            pStyle.self_closing = true;
            pStyle.attributes["w:val"] = "Normal";
            XML::Node bidi("w:bidi");
            bidi.self_closing = true;
            bidi.attributes["w:val"] = "0";
            XML::Node jc("w:jc");
            jc.self_closing = true;
            jc.attributes["w:val"] = "start";
            XML::Node rPr("w:rPr");

            pPr.add_child(pStyle);
            pPr.add_child(bidi);
            pPr.add_child(jc);
            pPr.add_child(rPr);
        }
        p.add_child(pPr);

        XML::Node r("w:r");
        {
            XML::Node rPr("w:rPr");
            r.add_child(rPr);
        }
        p.add_child(r);
    }
    return p;
}

//////////////////////
// Text definitions //
//////////////////////

inline DOCX::Text::Text(std::string set_text) {
    text = set_text;
}

////////////////////////////
// DOCX Utils definitions //
////////////////////////////

inline void DOCXUtils::mkdir(std::string dirpath) {
    std::filesystem::create_directory(dirpath);
}

inline void DOCXUtils::write_file(std::string fpath, std::string content) {
    std::ofstream ofs(fpath);
    ofs << content;
    ofs.close();
}

inline void DOCXUtils::delete_file_or_folder(std::string dirpath) {
    std::filesystem::remove(dirpath);
}

inline std::string DOCXUtils::content_types_file() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\"?> \
<Types \
	xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"> \
	<Default Extension=\"xml\" ContentType=\"application/xml\"/> \
	<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/> \
	<Default Extension=\"png\" ContentType=\"image/png\"/> \
	<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/> \
	<Override PartName=\"/_rels/.rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/> \
	<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/> \
	<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/> \
	<Override PartName=\"/word/_rels/document.xml.rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/> \
	<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/> \
	<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/> \
	<Override PartName=\"/word/fontTable.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml\"/> \
	<Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/> \
	<Override PartName=\"/word/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/> \
</Types>";
}

inline std::string DOCXUtils::dotrels_file() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\"?> \
<Relationships \
	xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"> \
	<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/> \
	<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/> \
	<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/> \
</Relationships>";
}

inline std::string DOCXUtils::app_file() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?> \
<Properties \
	xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" \
	xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"> \
	<Template></Template> \
	<TotalTime>0</TotalTime> \
	<Application>Adict</Application> \
	<AppVersion>1.0</AppVersion> \
	<Pages>1</Pages> \
	<Words>1</Words> \
	<Characters>1</Characters> \
	<CharactersWithSpaces>1</CharactersWithSpaces> \
	<Paragraphs>1</Paragraphs> \
</Properties>";
}

inline std::string DOCXUtils::core_file() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>  \
<cp:coreProperties \
	xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"  \
	xmlns:dc=\"http://purl.org/dc/elements/1.1/\" \
	xmlns:dcterms=\"http://purl.org/dc/terms/\" \
	xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" \
	xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"> \
	<dcterms:created xsi:type=\"dcterms:W3CDTF\">2025-05-23T22:33:19Z</dcterms:created> \
	<dc:creator></dc:creator> \
	<dc:description></dc:description> \
	<dc:language>en-US</dc:language> \
	<cp:lastModifiedBy></cp:lastModifiedBy> \
	<dcterms:modified xsi:type=\"dcterms:W3CDTF\">2025-05-23T22:34:12Z</dcterms:modified> \
	<cp:revision>2</cp:revision> \
	<dc:subject></dc:subject> \
	<dc:title></dc:title> \
</cp:coreProperties>";
}

inline std::string DOCXUtils::font_table_file() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?> \
<w:fonts \
	xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" \
	xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">  \
	<w:font w:name=\"Times New Roman\"> \
		<w:charset w:val=\"00\" w:characterSet=\"windows-1252\"/> \
		<w:family w:val=\"roman\"/> \
		<w:pitch w:val=\"variable\"/> \
	</w:font> \
	<w:font w:name=\"Symbol\"> \
		<w:charset w:val=\"02\"/> \
		<w:family w:val=\"roman\"/> \
		<w:pitch w:val=\"variable\"/> \
	</w:font> \
	<w:font w:name=\"Arial\"> \
		<w:charset w:val=\"00\" w:characterSet=\"windows-1252\"/> \
		<w:family w:val=\"swiss\"/> \
		<w:pitch w:val=\"variable\"/> \
	</w:font> \
	<w:font w:name=\"Liberation Serif\"> \
		<w:altName w:val=\"Times New Roman\"/> \
		<w:charset w:val=\"a2\" w:characterSet=\"windows-1254\"/> \
		<w:family w:val=\"roman\"/> \
		<w:pitch w:val=\"variable\"/> \
	</w:font> \
	<w:font w:name=\"Liberation Sans\"> \
		<w:altName w:val=\"Arial\"/> \
		<w:charset w:val=\"a2\" w:characterSet=\"windows-1254\"/> \
		<w:family w:val=\"swiss\"/> \
		<w:pitch w:val=\"variable\"/> \
	</w:font> \
</w:fonts>";
}

inline std::string DOCXUtils::settings_file() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?> \
<w:settings \
	xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"> \
	<w:zoom w:percent=\"100\"/> \
	<w:defaultTabStop w:val=\"709\"/> \
	<w:autoHyphenation w:val=\"true\"/> \
	<w:hyphenationZone w:val=\"0\"/> \
	<w:compat> \
		<w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"15\"/> \
		<w:compatSetting w:name=\"useWord2013TrackBottomHyphenation\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"1\"/> \
		<w:compatSetting w:name=\"allowHyphenationAtTrackBottom\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"1\"/> \
	</w:compat> \
</w:settings>";
}

inline std::string DOCXUtils::styles_file() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?> \
<w:styles \
	xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" \
	xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" \
	xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"w14\"> \
	<w:docDefaults> \
		<w:rPrDefault> \
			<w:rPr> \
				<w:rFonts w:ascii=\"Georgia\" w:hAnsi=\"Georgia\" w:eastAsia=\"NSimSun\" w:cs=\"Arial\"/> \
				<w:kern w:val=\"2\"/> \
				<w:sz w:val=\"24\"/> \
				<w:szCs w:val=\"24\"/> \
				<w:lang w:val=\"en-US\" w:eastAsia=\"zh-CN\" w:bidi=\"hi-IN\"/> \
			</w:rPr> \
		</w:rPrDefault> \
		<w:pPrDefault> \
			<w:pPr> \
				<w:widowControl/> \
				<w:suppressAutoHyphens w:val=\"true\"/> \
			</w:pPr> \
		</w:pPrDefault> \
	</w:docDefaults> \
	<w:style w:type=\"paragraph\" w:styleId=\"Normal\"> \
		<w:name w:val=\"Normal\"/> \
		<w:qFormat/> \
		<w:pPr> \
			<w:widowControl/> \
			<w:bidi w:val=\"0\"/> \
		</w:pPr> \
		<w:rPr> \
			<w:rFonts w:ascii=\"Georgia\" w:hAnsi=\"Georgia\" w:eastAsia=\"NSimSun\" w:cs=\"Arial\"/> \
			<w:color w:val=\"auto\"/> \
			<w:kern w:val=\"2\"/> \
			<w:sz w:val=\"24\"/> \
			<w:szCs w:val=\"24\"/> \
			<w:lang w:val=\"en-US\" w:eastAsia=\"zh-CN\" w:bidi=\"hi-IN\"/> \
		</w:rPr> \
	</w:style> \
	<w:style w:type=\"paragraph\" w:styleId=\"Heading\"> \
		<w:name w:val=\"Heading\"/> \
		<w:basedOn w:val=\"Normal\"/> \
		<w:next w:val=\"BodyText\"/> \
		<w:qFormat/> \
		<w:pPr> \
			<w:keepNext w:val=\"true\"/> \
			<w:spacing w:before=\"240\" w:after=\"120\"/> \
		</w:pPr> \
		<w:rPr> \
			<w:rFonts w:ascii=\"Georgia\" w:hAnsi=\"Georgia\" w:eastAsia=\"Microsoft YaHei\" w:cs=\"Arial\"/> \
			<w:sz w:val=\"28\"/> \
			<w:szCs w:val=\"28\"/> \
		</w:rPr> \
	</w:style> \
	<w:style w:type=\"paragraph\" w:styleId=\"BodyText\"> \
		<w:name w:val=\"Body Text\"/> \
		<w:basedOn w:val=\"Normal\"/> \
		<w:pPr> \
			<w:spacing w:lineRule=\"auto\" w:line=\"276\" w:before=\"0\" w:after=\"140\"/> \
		</w:pPr> \
		<w:rPr></w:rPr> \
	</w:style> \
	<w:style w:type=\"paragraph\" w:styleId=\"List\"> \
		<w:name w:val=\"List\"/> \
		<w:basedOn w:val=\"BodyText\"/> \
		<w:pPr></w:pPr> \
		<w:rPr> \
			<w:rFonts w:cs=\"Arial\"/> \
		</w:rPr> \
	</w:style> \
	<w:style w:type=\"paragraph\" w:styleId=\"Caption\"> \
		<w:name w:val=\"caption\"/> \
		<w:basedOn w:val=\"Normal\"/> \
		<w:qFormat/> \
		<w:pPr> \
			<w:suppressLineNumbers/> \
			<w:spacing w:before=\"120\" w:after=\"120\"/> \
		</w:pPr> \
		<w:rPr> \
			<w:rFonts w:cs=\"Arial\"/> \
			<w:i/> \
			<w:iCs/> \
			<w:sz w:val=\"24\"/> \
			<w:szCs w:val=\"24\"/> \
		</w:rPr> \
	</w:style> \
	<w:style w:type=\"paragraph\" w:styleId=\"Index\"> \
		<w:name w:val=\"Index\"/> \
		<w:basedOn w:val=\"Normal\"/> \
		<w:qFormat/> \
		<w:pPr> \
			<w:suppressLineNumbers/> \
		</w:pPr> \
		<w:rPr> \
			<w:rFonts w:cs=\"Arial\"/> \
		</w:rPr> \
	</w:style> \
</w:styles>";
}

inline std::string DOCXUtils::document_xml_rels_file() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\"?> \
<Relationships \
	xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"> \
	<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/> \
	<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable\" Target=\"fontTable.xml\"/> \
	<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/> \
	<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/> \
</Relationships>";
}

inline std::string DOCXUtils::theme1_file() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?> \
<a:theme \
	xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" \
	xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" name=\"Office\"> \
	<a:themeElements> \
		<a:clrScheme name=\"LibreOffice\"> \
			<a:dk1> \
				<a:srgbClr val=\"000000\"/> \
			</a:dk1> \
			<a:lt1> \
				<a:srgbClr val=\"ffffff\"/> \
			</a:lt1> \
			<a:dk2> \
				<a:srgbClr val=\"000000\"/> \
			</a:dk2> \
			<a:lt2> \
				<a:srgbClr val=\"ffffff\"/> \
			</a:lt2> \
			<a:accent1> \
				<a:srgbClr val=\"18a303\"/> \
			</a:accent1> \
			<a:accent2> \
				<a:srgbClr val=\"0369a3\"/> \
			</a:accent2> \
			<a:accent3> \
				<a:srgbClr val=\"a33e03\"/> \
			</a:accent3> \
			<a:accent4> \
				<a:srgbClr val=\"8e03a3\"/> \
			</a:accent4> \
			<a:accent5> \
				<a:srgbClr val=\"c99c00\"/> \
			</a:accent5> \
			<a:accent6> \
				<a:srgbClr val=\"c9211e\"/> \
			</a:accent6> \
			<a:hlink> \
				<a:srgbClr val=\"0000ee\"/> \
			</a:hlink> \
			<a:folHlink> \
				<a:srgbClr val=\"551a8b\"/> \
			</a:folHlink> \
		</a:clrScheme> \
		<a:fontScheme name=\"Office\"> \
			<a:majorFont> \
				<a:latin typeface=\"Arial\" pitchFamily=\"0\" charset=\"1\"/> \
				<a:ea typeface=\"DejaVu Sans\" pitchFamily=\"0\" charset=\"1\"/> \
				<a:cs typeface=\"DejaVu Sans\" pitchFamily=\"0\" charset=\"1\"/> \
			</a:majorFont> \
			<a:minorFont> \
				<a:latin typeface=\"Arial\" pitchFamily=\"0\" charset=\"1\"/> \
				<a:ea typeface=\"DejaVu Sans\" pitchFamily=\"0\" charset=\"1\"/> \
				<a:cs typeface=\"DejaVu Sans\" pitchFamily=\"0\" charset=\"1\"/> \
			</a:minorFont> \
		</a:fontScheme> \
		<a:fmtScheme> \
			<a:fillStyleLst> \
				<a:solidFill> \
					<a:schemeClr val=\"phClr\"></a:schemeClr> \
				</a:solidFill> \
				<a:solidFill> \
					<a:schemeClr val=\"phClr\"></a:schemeClr> \
				</a:solidFill> \
				<a:solidFill> \
					<a:schemeClr val=\"phClr\"></a:schemeClr> \
				</a:solidFill> \
			</a:fillStyleLst> \
			<a:lnStyleLst> \
				<a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"> \
					<a:prstDash val=\"solid\"/> \
					<a:miter/> \
				</a:ln> \
				<a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"> \
					<a:prstDash val=\"solid\"/> \
					<a:miter/> \
				</a:ln> \
				<a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"> \
					<a:prstDash val=\"solid\"/> \
					<a:miter/> \
				</a:ln> \
			</a:lnStyleLst> \
			<a:effectStyleLst> \
				<a:effectStyle> \
					<a:effectLst/> \
				</a:effectStyle> \
				<a:effectStyle> \
					<a:effectLst/> \
				</a:effectStyle> \
				<a:effectStyle> \
					<a:effectLst/> \
				</a:effectStyle> \
			</a:effectStyleLst> \
			<a:bgFillStyleLst> \
				<a:solidFill> \
					<a:schemeClr val=\"phClr\"></a:schemeClr> \
				</a:solidFill> \
				<a:solidFill> \
					<a:schemeClr val=\"phClr\"></a:schemeClr> \
				</a:solidFill> \
				<a:solidFill> \
					<a:schemeClr val=\"phClr\"></a:schemeClr> \
				</a:solidFill> \
			</a:bgFillStyleLst> \
		</a:fmtScheme> \
	</a:themeElements> \
</a:theme>";
}

#endif