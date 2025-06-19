#include "docx.hpp"

int main() {
    DOCX docx;

    DOCX::Paragraph p;
    p.add_text("hello world");
    p.add_space();
    p.add_italic_text("this is italic");
    docx.add_paragraph(p);

    DOCX::Paragraph p2;
    p2.add_bold_text("this is bold");
    p2.add_space();
    p2.add_text("this is not bold");
    docx.add_paragraph(p2);

    docx.add_empty_line();

    docx.add_paragraph(p);
    
    docx.print();
    docx.save("my_document.docx");
    return 0;
}