from Part1 import path_and_sheets, process_sheet, export_json, generate_pdf_from_json, pdf_from_json
from Part2 import avg_sum_of_each_sheet
from fastapi import FastAPI, UploadFile, File, HTTPException, Request
import os
import matplotlib.pyplot as plt
import json
from reportlab.pdfgen import canvas
from fastapi.responses import JSONResponse


app = FastAPI()


@app.get("/")
def root():
    return {"message": "Hello World"}


# /docs - swagger

@app.post("/upload_file/")
async def upload_file(file: UploadFile = File(...)):
    try:
        return path_and_sheets(file)
    except:
        raise HTTPException(status_code=400, detail="An error occurred while processing the file")


@app.post("/receive_data/")
async def receive_data(request_data: dict):
    try:
        return export_json(request_data)
    except Exception:
        error_message = "An error occurred while generating the report."
        raise HTTPException(status_code=500, detail=error_message)


@app.post("/create_pdf")
async def create_pdf_from_json(request: Request):
    try:
        return pdf_from_json(request)
    except:
        raise HTTPException(status_code=500, detail="Json pdf did not work")


@app.post("/all_data/")
async def all_data(file: UploadFile = File(...)):
    try:
        os.makedirs('files', exist_ok=True)
        excel_path = f"files/{file.filename}"

        with open(excel_path, "wb") as f:
            f.write(file.file.read())
        data = {'name': os.path.basename(excel_path), 'sheets':avg_sum_of_each_sheet(excel_path) }
        pdf_file = "all_data.pdf"
        c = canvas.Canvas(pdf_file)
        c.drawString(10, 800, text=json.dumps(data, indent=4))

        # Create and save the matplotlib chart
        sheet_names = [data['sheets']['name'] for sheet in data['sheets']]
        sum_values = [data['sheets']['sum'] for sheet in data['sheets']]
        avg_values = [data['sheets']['avg'] for sheet in data['sheets']]

        plt.figure(figsize=(10, 5))

        plt.bar(sheet_names, sum_values, color='skyblue', label='Sum')
        plt.bar(sheet_names, avg_values, color='salmon', label='Average')

        plt.xlabel('Sheet Names')
        plt.ylabel('Values')
        plt.title('Sum and Average for each Sheet')
        plt.legend()
        chart_file = "chart.png"
        plt.savefig(chart_file)

        # Draw the chart in the PDF
        c.drawImage(chart_file, 10, 500, width=100, height=100)

        # Save and close the PDF
        c.save()
        plt.close()
        # return report_data
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})


if __name__ == "__main__":
    import uvicorn

    uvicorn.run('main:app', host="localhost", port=7000, reload=True)
