from gspread.urls import SPREADSHEETS_API_V4_BASE_URL


def insert_note(worksheet, label, note):
    """
    Insert note ito the google worksheet for a certain cell.

    Compatible with gspread.
    """
    spreadsheet_id = worksheet.spreadsheet.id
    worksheet_id = worksheet.id

    row, col = tuple(label)      # [0, 0] is A1

    url = f"{SPREADSHEETS_API_V4_BASE_URL}/{spreadsheet_id}:batchUpdate"
    payload = {
        "requests": [
            {
                "updateCells": {
                    "range": {
                        "sheetId": worksheet_id,
                        "startRowIndex": row,
                        "endRowIndex": row + 1,
                        "startColumnIndex": col,
                        "endColumnIndex": col + 1
                    },
                    "rows": [
                        {
                            "values": [
                                {
                                    "note": note
                                }
                            ]
                        }
                    ],
                    "fields": "note"
                }
            }
        ]
    }
    worksheet.spreadsheet.client.request("post", url, json=payload)
