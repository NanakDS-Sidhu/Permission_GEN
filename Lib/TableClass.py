from docx import Document
from docx2pdf import convert

class TableGenerator:
    def __init__(self, template_path,people,temp,table_number):
        self.template_path = template_path
        self.document = Document(template_path)
        self.table = self.document.tables[table_number]
        self.people=people
        self.temporary_file=temp

    def insert_people_into_table(self):
        people=self.people
        count = 0
        for row in self.table.rows:
            for cell in row.cells:
                if "{Sno}" in cell.text:
                    cell.text = cell.text.replace("{Sno}", str(count) if people[count-1] else "")
                elif "{name}" in cell.text:
                    cell.text = cell.text.replace("{name}", people[count-1]["name"] if people[count-1] else "")
                elif "{SID}" in cell.text:
                    cell.text = cell.text.replace("{SID}", people[count-1]["SID"] if people[count-1] else "")
                elif "{hostel}" in cell.text:
                    cell.text = cell.text.replace("{hostel}", people[count-1]["Hostel"] if people[count-1] else "")
            count += 1

    def add_rows(self):
        num_rows=len(self.people)
        for _ in range(num_rows):
            self.table.add_row()

    def fill_rows(self):
        people=self.people
        n = len(people)
        cols = 4
        content = ["{Sno}", "{name}", "{SID}", "{hostel}"]
        for j in range(cols):
            for i in range(1, n+1):
                self.table.cell(i, j).text = content[j]

    def save_as_docx(self, output_path):
        self.document.save(output_path)

    def run(self):
        self.add_rows()
        self.fill_rows()
        self.insert_people_into_table()
        self.save_as_docx(self.temporary_file)
    

