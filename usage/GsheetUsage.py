import Lib.GSheetsClass as gs

def main():
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
    SPREADSHEET_ID = "1uGdpcN231nAnUJ3yPlwFexUnCGi1f-UPhitgWnjvpR0"

    google_sheets_handler = gs.GoogleSheetsHandler(SCOPES, SPREADSHEET_ID)
    data = google_sheets_handler.fetch_data_from_spreadsheet("Sheet1")

    if data:
        for row in data:
            print(row)

if __name__ == "__main__":
    main()