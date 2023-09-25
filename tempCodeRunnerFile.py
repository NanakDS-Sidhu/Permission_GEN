ator = tc.TableGenerator(template_path,people,temporary_path)
    letter_generator.run()
    convert("temp.docx", 'output.pdf')   
