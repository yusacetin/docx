#ifndef DOCX_HPP
#define DOCX_HPP

#include "../xml/xml.hpp"

#include <string>
#include <cstdlib>

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
    
private:
    std::vector<DOCX::Paragraph> paragraphs;
    XML::Node get();

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
    XML::Node root = get();
    root.save("docx_root/word/document.xml");
    std::string save_cmd = "( cd docx_root && zip -r ../" + fname + " . > /dev/null 2>&1 )";
    system(save_cmd.c_str());
    std::cout << "Saved as " << fname << newl;
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

#endif