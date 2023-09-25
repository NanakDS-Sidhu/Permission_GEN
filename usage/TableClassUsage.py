from docx2pdf import convert
from Lib import TableClass as tc

def main():
    template_path = 'template.docx'
    temporary_path="temp.docx"
    people = [
        {"name": "John Doe", "SID": "12345", "Hostel": "A"},
        {"name": "Jane Smith", "SID": "67890", "Hostel": "B"},
        {"name": "Jacob Ladders", "SID": "67891", "Hostel": "B"},
        # Add more people as needed
    ]

    letter_generator = tc.TableGenerator(template_path,people,temporary_path)
    letter_generator.run()
    convert("temp.docx", 'output.pdf')

if __name__=="__main__":
    main()

