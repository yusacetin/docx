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
    void add_empty_line(size_t count = 1, size_t font_size = 0);
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
	size_t empty_line_size = 0; // ignored if not empty

	void add_text(Text t);
	void add_text(std::string text_str);
    void add_formatted_text(Text t);
    void add_plain_text(std::string text_str);
    void add_space(size_t count = 1, size_t font_size = 0); // 0 to follow global setting
    void add_bold_text(std::string text_str);
    void add_italic_text(std::string text_str);
    void add_underlined_text(std::string text_str);
    void add_struckthrough_text(std::string text_str);
    XML::Node get();
    
    static XML::Node empty_line_node(size_t set_empty_line_size = 0); // 0 to follow global setting

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

inline void DOCX::add_empty_line(size_t count, size_t font_size) {
    for (size_t i = 0; i < count; i++) {
        DOCX::Paragraph p;
        p.empty = true;
		p.empty_line_size = font_size;
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
            body.add_child(DOCX::Paragraph::empty_line_node(cur_p.empty_line_size));
        } else {
            body.add_child(paragraphs.at(i).get());
        }
    }

	// Add paragraph properties

	XML::Node sectPr("w:sectPr");
	{
		XML::Node pgMar("w:pgMar");
		pgMar.self_closing = true;
		pgMar.attributes["w:top"] = "720";
		pgMar.attributes["w:right"] = "720";
		pgMar.attributes["w:bottom"] = "720";
		pgMar.attributes["w:left"] = "720";
		pgMar.attributes["w:header"] = "360";
		pgMar.attributes["w:footer"] = "360";
		pgMar.attributes["w:gutter"] = "0";
		sectPr.add_child(pgMar);
	}
	body.add_child(sectPr);

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

inline void DOCX::Paragraph::add_text(Text t) {
	contents.push_back(t);
}

inline void DOCX::Paragraph::add_text(std::string text_str) {
	Text t(text_str);
	contents.push_back(t);
}

inline void DOCX::Paragraph::add_formatted_text(Text t) { // same as add_text(Text)
    contents.push_back(t);
}

inline void DOCX::Paragraph::add_plain_text(std::string text_str) { // same as add_text(std::string)
    Text t(text_str);
    contents.push_back(t);
}

inline void DOCX::Paragraph::add_space(size_t count, size_t font_size) {
    std::string spaces;
    for (size_t i = 0; i < count; i++) {
        spaces += " ";
    }

    Text t(spaces);
    t.preserve_space = true;
	if (font_size > 0) {
		t.size = font_size;
	}
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

inline XML::Node DOCX::Paragraph::empty_line_node(size_t set_empty_line_size) {
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
			{
				if (set_empty_line_size > 0) {
					XML::Node sz("w:sz");
					sz.self_closing = true;
					sz.attributes["w:val"] = std::to_string(set_empty_line_size * 2); // because half points
					rPr.add_child(sz);

					XML::Node szCs("w:szCs");
					szCs.self_closing = true;
					szCs.attributes["w:val"] = std::to_string(set_empty_line_size * 2); // because half points
					rPr.add_child(szCs);
				}
			}

            pPr.add_child(pStyle);
            pPr.add_child(bidi);
            pPr.add_child(jc);
            pPr.add_child(rPr);
        }
        p.add_child(pPr);

        XML::Node r("w:r");
        {
            XML::Node rPr("w:rPr");
			if (set_empty_line_size > 0) {
				XML::Node sz("w:sz");
				sz.self_closing = true;
				sz.attributes["w:val"] = std::to_string(set_empty_line_size * 2); // because half points
				rPr.add_child(sz);

				XML::Node szCs("w:szCs");
				szCs.self_closing = true;
				szCs.attributes["w:val"] = std::to_string(set_empty_line_size * 2); // because half points
				rPr.add_child(szCs);
			}
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
	XML::Node types("Types");
	types.attributes["xmlns"] = "http://schemas.openxmlformats.org/package/2006/content-types";
	{
		XML::Node def1("Default");
		def1.attributes["Extension"] = "xml";
		def1.attributes["ContentType"] = "application/xml";
		def1.self_closing = true;
		types.add_child(def1);

		XML::Node def2("Default");
		def2.attributes["Extension"] = "rels";
		def2.attributes["ContentType"] = "application/vnd.openxmlformats-package.relationships+xml";
		def2.self_closing = true;
		types.add_child(def2);

		XML::Node def3("Default");
		def3.attributes["Extension"] = "png";
		def3.attributes["ContentType"] = "image/png";
		def3.self_closing = true;
		types.add_child(def3);

		XML::Node def4("Default");
		def4.attributes["Extension"] = "jpeg";
		def4.attributes["ContentType"] = "image/jpeg";
		def4.self_closing = true;
		types.add_child(def4);

		XML::Node over1("Override");
		over1.attributes["PartName"] = "/_rels/.rels";
		over1.attributes["ContentType"] = "application/vnd.openxmlformats-package.relationships+xml";
		over1.self_closing = true;
		types.add_child(over1);

		XML::Node over2("Override");
		over2.attributes["PartName"] = "/docProps/core.xml";
		over2.attributes["ContentType"] = "application/vnd.openxmlformats-package.core-properties+xml";
		over2.self_closing = true;
		types.add_child(over2);

		XML::Node over3("Override");
		over3.attributes["PartName"] = "/docProps/app.xml";
		over3.attributes["ContentType"] = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
		over3.self_closing = true;
		types.add_child(over3);

		XML::Node over4("Override");
		over4.attributes["PartName"] = "/word/_rels/document.xml.rels";
		over4.attributes["ContentType"] = "application/vnd.openxmlformats-package.relationships+xml";
		over4.self_closing = true;
		types.add_child(over4);

		XML::Node over5("Override");
		over5.attributes["PartName"] = "/word/document.xml";
		over5.attributes["ContentType"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
		over5.self_closing = true;
		types.add_child(over5);

		XML::Node over6("Override");
		over6.attributes["PartName"] = "/word/styles.xml";
		over6.attributes["ContentType"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
		over6.self_closing = true;
		types.add_child(over6);

		XML::Node over7("Override");
		over7.attributes["PartName"] = "/word/fontTable.xml";
		over7.attributes["ContentType"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
		over7.self_closing = true;
		types.add_child(over7);

		XML::Node over8("Override");
		over8.attributes["PartName"] = "/word/settings.xml";
		over8.attributes["ContentType"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
		over8.self_closing = true;
		types.add_child(over8);
		
		XML::Node over9("Override");
		over9.attributes["PartName"] = "/word/theme/theme1.xml";
		over9.attributes["ContentType"] = "application/vnd.openxmlformats-officedocument.theme+xml";
		over9.self_closing = true;
		types.add_child(over9);
	}
	return types.get_string();
}

inline std::string DOCXUtils::dotrels_file() {
	XML::Node rels("Relationships");
	rels.attributes["xmlns"] = "http://schemas.openxmlformats.org/package/2006/relationships";
	{
		XML::Node rel1("Relationship");
		rel1.attributes["Id"] = "rId1";
		rel1.attributes["Type"] = "http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties";
		rel1.attributes["Target"] = "docProps/core.xml";
		rels.add_child(rel1);

		XML::Node rel2("Relationship");
		rel2.attributes["Id"] = "rId2";
		rel2.attributes["Type"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
		rel2.attributes["Target"] = "docProps/app.xml";
		rels.add_child(rel2);

		XML::Node rel3("Relationship");
		rel3.attributes["Id"] = "rId3";
		rel3.attributes["Type"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
		rel3.attributes["Target"] = "word/document.xml";
		rels.add_child(rel3);
	}
	return rels.get_string();
}

inline std::string DOCXUtils::app_file() {
    XML::Node props("Properties");
	props.attributes["xmlns"] = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
	props.attributes["xmlns:vt"] = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
	{
		XML::Node temp("Template");
		props.add_child(temp);

		XML::Node total_time("TotalTime");
		total_time.content = "0";
		props.add_child(total_time);

		XML::Node application("Application");
		application.content = "Adict";
		props.add_child(application);

		XML::Node app_version("AppVersion");
		app_version.content = "1.0";
		props.add_child(app_version);

		XML::Node pages("Pages");
		pages.content = "1";
		props.add_child(pages);

		XML::Node words("Words");
		words.content = "1";
		props.add_child(words);

		XML::Node chars("Characters");
		chars.content = "1";
		props.add_child(chars);

		XML::Node charsws("CharactersWithSpaces");
		charsws.content = "1";
		props.add_child(charsws);

		XML::Node pars("Paragraphs");
		pars.content = "1";
		props.add_child(pars);
	}
	return props.get_string();
}

inline std::string DOCXUtils::core_file() {
    XML::Node cp("cp:coreProperties");
	cp.attributes["xmlns:cp"] = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
	cp.attributes["xmlns:dc"] = "http://purl.org/dc/elements/1.1/";
	cp.attributes["xmlns:dcterms"] = "http://purl.org/dc/terms/";
	cp.attributes["xmlns:dcmitype"] = "http://purl.org/dc/dcmitype/";
	cp.attributes["xmlns:xsi"] = "http://www.w3.org/2001/XMLSchema-instance";
	{
		XML::Node created("dcterms:created");
		created.attributes["xsi:type"] = "dcterms:W3CDTF";
		created.content = "2025-06-01T12:00:00Z"; // TODO make dynamic
		cp.add_child(created);

		XML::Node creator("dc:creator");
		cp.add_child(creator);

		XML::Node desc("dc:description");
		cp.add_child(desc);

		XML::Node lang("dc:language");
		lang.content = "en-US";
		cp.add_child(lang);

		XML::Node lastmod("cp:lastModifiedBy");
		cp.add_child(lastmod);

		XML::Node modified("dcterms:modified");
		modified.attributes["xsi:type"] = "dcterms:W3CDTF";
		modified.content = "2025-06-01T12:00:00Z"; // TODO make dynamic
		cp.add_child(modified);

		XML::Node rev("cp:revision");
		rev.content = "2";
		cp.add_child(rev);

		XML::Node subject("dc:subject");
		cp.add_child(subject);

		XML::Node title("dc:title");
		cp.add_child(title);
	}
	return cp.get_string();
}

inline std::string DOCXUtils::font_table_file() {
    XML::Node fonts("w:fonts");
	fonts.attributes["xmlns:w"] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
	fonts.attributes["xmlns:r"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
	{
		XML::Node font("w:font");
		font.attributes["w:name"] = "Georgia";
		{
			XML::Node charset("w:charset");
			charset.attributes["w:val"] = "00";
			charset.self_closing = true;
			font.add_child(charset);

			XML::Node family("w:family");
			family.attributes["w:val"] = "roman";
			family.self_closing = true;
			font.add_child(family);

			XML::Node pitch("w:pitch");
			pitch.attributes["w:val"] = "variable";
			pitch.self_closing = true;
			font.add_child(pitch);
		}
		fonts.add_child(font);
	}
	return fonts.get_string();
}

inline std::string DOCXUtils::settings_file() {
    XML::Node settings("w:settings");
	settings.attributes["xmlns:w"] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
	{
		XML::Node compat("w:compat");
		{
			XML::Node compatset("w:compatSetting");
			compatset.attributes["w:name"] = "compatibilityMode";
			compatset.attributes["w:uri"] = "http://schemas.microsoft.com/office/word";
			compatset.attributes["w:val"] = "15";
			compatset.self_closing = true;
			compat.add_child(compatset);
		}
		settings.add_child(compat);
	}
	return settings.get_string();
}

inline std::string DOCXUtils::styles_file() {
    XML::Node styles("w:styles");
	styles.attributes["xmlns:w"] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
	styles.attributes["xmlns:w14"] = "http://schemas.microsoft.com/office/word/2010/wordml";
	styles.attributes["xmlns:mc"] = "http://schemas.openxmlformats.org/markup-compatibility/2006";
	styles.attributes["mc:Ignorable"] = "w14";
	{
		XML::Node docdefs("w:docDefaults");
		{
			XML::Node rprdef1("w:rPrDefault");
			{
				XML::Node rpr("w:rPr");
				{
					XML::Node rfonts("w:rFonts");
					rfonts.attributes["w:ascii"] = "Georgia";
					rfonts.attributes["w:hAnsi"] = "Georgia";
					rfonts.attributes["w:eastAsia"] = "Georgia";
					rfonts.attributes["w:cs"] = "Georgia";
					rfonts.self_closing = true;
					rpr.add_child(rfonts);

					XML::Node kern("w:kern");
					kern.attributes["w:val"] = "2";
					kern.self_closing = true;
					rpr.add_child(kern);

					XML::Node sz("w:sz");
					sz.attributes["w:val"] = "24";
					sz.self_closing = true;
					rpr.add_child(sz);

					XML::Node szcs("w:szCs");
					szcs.attributes["w:val"] = "24";
					szcs.self_closing = true;
					rpr.add_child(szcs);

					XML::Node lang("w:lang");
					lang.attributes["w:val"] = "en-US";
					lang.self_closing = true;
					rpr.add_child(lang);
				}
				rprdef1.add_child(rpr);
			}
			docdefs.add_child(rprdef1);

			XML::Node pprdef("w:pPrDefault");
			{
				XML::Node ppr("w:pPr");
				{
					XML::Node win("w:windowControl");
					win.self_closing = true;
					ppr.add_child(win);

					XML::Node sup("w:suppressAutoHyphens");
					sup.attributes["w:val"] = "true";
					sup.self_closing = true;
					ppr.add_child(sup);
				}
				pprdef.add_child(ppr);
			}
			docdefs.add_child(pprdef);
		}
		styles.add_child(docdefs);

		XML::Node style("w:style");
		style.attributes["w:type"] = "paragraph";
		style.attributes["w:default"] = "1";
		style.attributes["w:styleId"] = "Normal";
		{
			XML::Node name("w:name");
			name.attributes["w:val"] = "Normal";
			name.self_closing = true;
			style.add_child(name);

			XML::Node qformat("w:qFormat");
			qformat.self_closing = true;
			style.add_child(qformat);

			XML::Node rpr("w:rPr");
			{
				XML::Node rfonts("w:rFonts");
				rfonts.attributes["w:ascii"] = "Georgia";
				rfonts.attributes["w:hAnsi"] = "Georgia";
				rfonts.attributes["w:eastAsia"] = "Georgia";
				rfonts.attributes["w:cs"] = "Georgia";
				rfonts.self_closing = true;
				rpr.add_child(rfonts);
			}
			style.add_child(rpr);
		}
		styles.add_child(style);
	}
	return styles.get_string();
}

inline std::string DOCXUtils::document_xml_rels_file() {
	XML::Node rels("Relationships");
	rels.attributes["xmlns"] = "http://schemas.openxmlformats.org/package/2006/relationships";
	{
		XML::Node rel1("Relationship");
		rel1.attributes["Id"] = "rId1";
		rel1.attributes["Type"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
		rel1.attributes["Target"] = "styles.xml";
		rels.add_child(rel1);

		XML::Node rel2("Relationship");
		rel2.attributes["Id"] = "rId2";
		rel2.attributes["Type"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
		rel2.attributes["Target"] = "fontTable.xml";
		rels.add_child(rel2);

		XML::Node rel3("Relationship");
		rel3.attributes["Id"] = "rId3";
		rel3.attributes["Type"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
		rel3.attributes["Target"] = "settings.xml";
		rels.add_child(rel3);

		XML::Node rel4("Relationship");
		rel4.attributes["Id"] = "rId4";
		rel4.attributes["Type"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
		rel4.attributes["Target"] = "theme/theme1.xml";
		rels.add_child(rel4);
	}
	return rels.get_string();
}

inline std::string DOCXUtils::theme1_file() {
    XML::Node theme1("a:theme");
	theme1.attributes["xmlns:a"] = "http://schemas.openxmlformats.org/drawingml/2006/main";
	theme1.attributes["xmlns:r"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
	theme1.attributes["name"] = "MinimalTheme";
	{
		XML::Node theme_elems("a:themeElements");
		{
			XML::Node clr_scheme("a:clrScheme");
			clr_scheme.attributes["name"] = "MinimalColors";
			{
				XML::Node dk1("a:dk1");
				{
					XML::Node srgb("a:srgbClr");
					srgb.attributes["val"] = "000000";
					srgb.self_closing = true;
					dk1.add_child(srgb);
				}
				clr_scheme.add_child(dk1);

				XML::Node lt1("a:lt1");
				{
					XML::Node srgb("a:srgbClr");
					srgb.attributes["val"] = "FFFFFF";
					srgb.self_closing = true;
					lt1.add_child(srgb);
				}
				clr_scheme.add_child(lt1);

				XML::Node dk2("a:dk2");
				{
					XML::Node srgb("a:srgbClr");
					srgb.attributes["val"] = "000000";
					srgb.self_closing = true;
					dk2.add_child(srgb);
				}
				clr_scheme.add_child(dk2);

				XML::Node lt2("a:lt2");
				{
					XML::Node srgb("a:srgbClr");
					srgb.attributes["val"] = "FFFFFF";
					srgb.self_closing = true;
					lt2.add_child(srgb);
				}
				clr_scheme.add_child(lt2);

				XML::Node acc1("a:accent1");
				{
					XML::Node srgb("a:srgbClr");
					srgb.attributes["val"] = "808080";
					srgb.self_closing = true;
					acc1.add_child(srgb);
				}
				clr_scheme.add_child(acc1);

				XML::Node acc2("a:accent2");
				{
					XML::Node srgb("a:srgbClr");
					srgb.attributes["val"] = "808080";
					srgb.self_closing = true;
					acc2.add_child(srgb);
				}
				clr_scheme.add_child(acc2);

				XML::Node acc3("a:accent3");
				{
					XML::Node srgb("a:srgbClr");
					srgb.attributes["val"] = "808080";
					srgb.self_closing = true;
					acc3.add_child(srgb);
				}
				clr_scheme.add_child(acc3);

				XML::Node acc4("a:accent4");
				{
					XML::Node srgb("a:srgbClr");
					srgb.attributes["val"] = "808080";
					srgb.self_closing = true;
					acc4.add_child(srgb);
				}
				clr_scheme.add_child(acc4);

				XML::Node acc5("a:accent5");
				{
					XML::Node srgb("a:srgbClr");
					srgb.attributes["val"] = "808080";
					srgb.self_closing = true;
					acc5.add_child(srgb);
				}
				clr_scheme.add_child(acc5);

				XML::Node acc6("a:accent6");
				{
					XML::Node srgb("a:srgbClr");
					srgb.attributes["val"] = "808080";
					srgb.self_closing = true;
					acc6.add_child(srgb);
				}
				clr_scheme.add_child(acc6);

				XML::Node hlink("a:hlink");
				{
					XML::Node srgb("a:srgbClr");
					srgb.attributes["val"] = "0000FF";
					srgb.self_closing = true;
					hlink.add_child(srgb);
				}
				clr_scheme.add_child(hlink);

				XML::Node fol("a:folHlink");
				{
					XML::Node srgb("a:srgbClr");
					srgb.attributes["val"] = "808080";
					srgb.self_closing = true;
					fol.add_child(srgb);
				}
				clr_scheme.add_child(fol);
			}
			theme_elems.add_child(clr_scheme);

			XML::Node font_scheme("a:fontScheme");
			font_scheme.attributes["name"] = "GeorgiaFont";
			{
				XML::Node major("a:majorFont");
				{
					XML::Node latin("a:latin");
					latin.attributes["typeface"] = "Georgia";
					latin.self_closing = true;
					major.add_child(latin);

					XML::Node ea("a:ea");
					ea.attributes["typeface"] = "Georgia";
					ea.self_closing = true;
					major.add_child(ea);

					XML::Node cs("a:cs");
					cs.attributes["typeface"] = "Georgia";
					cs.self_closing = true;
					major.add_child(cs);
				}
				font_scheme.add_child(major);

				XML::Node minor("a:minorFont");
				{
					XML::Node latin("a:latin");
					latin.attributes["typeface"] = "Georgia";
					latin.self_closing = true;
					minor.add_child(latin);

					XML::Node ea("a:ea");
					ea.attributes["typeface"] = "Georgia";
					ea.self_closing = true;
					minor.add_child(ea);

					XML::Node cs("a:cs");
					cs.attributes["typeface"] = "Georgia";
					cs.self_closing = true;
					minor.add_child(cs);
				}
				font_scheme.add_child(minor);
			}
			theme_elems.add_child(font_scheme);

			XML::Node fmt("a:fmtScheme");
			fmt.attributes["name"] = "MinimalFormat";
			{
				XML::Node fill("a:fillStyleLst");
				{
					XML::Node solid("a:solidFill");
					{
						XML::Node scheme("a:schemeClr");
						scheme.attributes["val"] = "phClr";
						scheme.self_closing = true;
						solid.add_child(scheme);
					}
					fill.add_child(solid);
				}
				fmt.add_child(fill);

				XML::Node ln("a:lnStyleLst");
				{
					XML::Node l("a:ln");
					l.attributes["w"] = "9525";
					{
						XML::Node solid("a:solidFill");
						{
							XML::Node scheme("a:schemeClr");
							scheme.attributes["val"] = "phClr";
							scheme.self_closing = true;
							solid.add_child(scheme);
						}
						l.add_child(solid);
					}
					ln.add_child(l);
				}
				fmt.add_child(ln);

				XML::Node effect("a:effectStyleLst");
				{
					XML::Node efs("a:effectStyle");
					{
						XML::Node efsl("a:effectLst");
						efsl.self_closing = true;
						efs.add_child(efsl);
					}
					effect.add_child(efs);
				}
				fmt.add_child(effect);

				XML::Node bg("a:bgFillStyleLst");
				{
					XML::Node solid("a:solidFill");
					{
						XML::Node scheme("a:schemeClr");
						scheme.attributes["val"] = "phClr";
						scheme.self_closing = true;
						solid.add_child(scheme);
					}
					bg.add_child(solid);
				}
				fmt.add_child(bg);
			}
			theme_elems.add_child(fmt);
		}
		theme1.add_child(theme_elems);
	}
	return theme1.get_string();
}

#endif