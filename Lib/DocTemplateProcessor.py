from docx import Document

class DocumentTemplateProcessor:
    def __init__(self, template_path):
        self.template_path = template_path
        self.document = Document(template_path)

    def replace_text_in_paragraphs(self, replacements):
        for paragraph in self.document.paragraphs:
            for key, value in replacements.items():
                if key in paragraph.text:
                    inline = paragraph.runs
                    for i in range(len(inline)):
                        if key in inline[i].text:
                            # print(inline[i].text,key,"wee")
                            inline[i].text = inline[i].text.replace(key, value)

    def save_modified_document(self, output_path):
        self.document.save(output_path)


    def process_document(self, replacements, output_docx_path):
        self.replace_text_in_paragraphs(replacements)
        self.save_modified_document(output_docx_path)

