from docx2pdf import convert
from Lib import TableClass as tc
import Lib.GSheetsClass as gs
import Lib.DocTemplateProcessor as dp

def create_dict(data):
    people=[]
    i=0
    for row in data:
        if i==0:
            i+=1
            continue
        temp={}
        attributes=data[0]
        for cell in range(len(row)):
            temp[attributes[cell]]=row[cell]
        people.append(temp)
        i+=1
    return people

def main():
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
    SPREADSHEET_ID = "1rCs_k4gq6mcNrjm9vTj8Ic6cOyRTka4-3QrepRD5gwY"

    google_sheets_handler = gs.GoogleSheetsHandler(SCOPES, SPREADSHEET_ID)
    data = google_sheets_handler.fetch_data_from_spreadsheet("Sheet1")
    # print(data)

    template_path = 'template.docx'
    temporary_path="temp.docx"
    
    people=create_dict(data)

    replacements = {
        "secy_SID": "2010101",
        "secy_name": "Agrim Arya",
        "start_time": "8pm",
        "end_time": "10pm",
        "start_date": "8th October 2023",
        "end_date": "9th October 2023"
    }

    processor = dp.DocumentTemplateProcessor(template_path)
    processor.process_document(replacements, "temp.docx")

    # print(people)
    table_number=0
    table_generator = tc.TableGenerator("temp.docx",people,temporary_path,table_number)
    table_generator.run()
    convert("temp.docx", 'output.pdf')   

if __name__ == "__main__":
    main()