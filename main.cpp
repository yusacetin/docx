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

    DOCX::Paragraph title;
    DOCX::Text t("Welcome to My Document");
    t.size = 24;
    title.add_text(t);
    title.align = DOCX::Paragraph::alignment::CENTER;
    docx.add_paragraph(title);

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

    docx.add_empty_line(1, 20);

    DOCX::Paragraph p10;
    DOCX::Text t10("There's only one space above this line but it has a huge font size");
    p10.add_text(t10);
    docx.add_paragraph(p10);

    docx.add_empty_line(2, 10);

    DOCX::Paragraph p11;
    DOCX::Text t11("There are two empty lines above this line but they are half the size of the previous empty line so the space should appear the same");
    p11.add_text(t11);
    docx.add_paragraph(p11);

    docx.add_empty_line();

    DOCX::Paragraph p12;
    DOCX::Text t12("Look at this");
    p12.add_text(t12);
    p12.add_space();
    p12.add_text("space, and");
    p12.add_space();
    DOCX::Text t122("now, look at this");
    t122.size = 20;
    p12.add_text(t122);
    p12.add_space();
    DOCX::Text t123("space and this");
    t123.size = 20;
    p12.add_text(t123);
    p12.add_space(1, 20);
    DOCX::Text t124("space");
    t124.size = 20;
    p12.add_text(t124);
    docx.add_paragraph(p12);

    docx.add_empty_line();

    DOCX::Paragraph p13;
    p13.add_text("This text is right aligned");
    p13.align = DOCX::Paragraph::alignment::RIGHT;
    docx.add_paragraph(p13);

    docx.add_empty_line();

    DOCX::Paragraph p14;
    p14.add_text("And this text is justified but in order for the justification to be observable it needs to span multiple lines so I'll put some lorem ipsum here lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");
    p14.align = DOCX::Paragraph::alignment::JUSTIFIED;
    docx.add_paragraph(p14);
    
    docx.save("my_document.docx");
    return 0;
}