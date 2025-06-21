#include "docx.hpp"

int main() {
    DOCX docx;

    DOCX::Paragraph p;
    p.add_plain_text("hello world");
    p.add_space();
    p.add_italic_text("this is italic");
    docx.add_paragraph(p);

    DOCX::Paragraph p2;
    p2.add_bold_text("this is bold");
    p2.add_space();
    p2.add_plain_text("this is not bold");
    docx.add_paragraph(p2);

    docx.add_empty_line();

    docx.add_paragraph(p);

    DOCX::Paragraph p3;
    DOCX::Text t3("This is both bold and italic!");
    t3.bold = true;
    t3.italic = true;
    p3.add_formatted_text(t3);
    docx.add_paragraph(p3);

    docx.add_empty_line();

    DOCX::Paragraph p4;
    DOCX::Text t4("This text is just plain");
    p4.add_formatted_text(t4);
    docx.add_paragraph(p4);

    docx.add_empty_line();

    DOCX::Paragraph p5;
    p5.add_underlined_text("hello underline");
    p5.add_space();
    p5.add_plain_text("and");
    p5.add_space();
    p5.add_struckthrough_text("strikethrough");
    docx.add_paragraph(p5);

    docx.add_empty_line();

    DOCX::Paragraph p6;
    DOCX::Text t6("Wow,");
    p6.add_formatted_text(t6);
    p6.add_space();
    DOCX::Text t62("these words");
    t62.bold = true;
    t62.italic = true;
    t62.underline = true;
    t62.strikethrough = true;
    p6.add_formatted_text(t62);
    p6.add_space();
    DOCX::Text t63("have all the formatting!");
    p6.add_formatted_text(t63);
    docx.add_paragraph(p6);

    docx.add_empty_line();

    DOCX::Paragraph p7;
    p7.add_plain_text("Look at all");
    p7.add_space(6);
    p7.add_plain_text("these");
    p7.add_space(4);
    p7.add_plain_text("spaces");
    docx.add_paragraph(p7);

    docx.add_empty_line(5);

    DOCX::Paragraph p8;
    p8.add_plain_text("There are 5 empty lines above this line!");
    docx.add_paragraph(p8);

    DOCX::Paragraph p9;
    DOCX::Text t9("this text is smol");
    t9.size = 9;
    p9.add_formatted_text(t9);
    p9.add_space();
    DOCX::Text t92("this text is HUGE!");
    t92.size = 20;
    p9.add_formatted_text(t92);
    docx.add_paragraph(p9);
    
    docx.save("my_document.docx");
    return 0;
}