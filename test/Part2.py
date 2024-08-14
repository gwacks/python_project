import openpyxl as xl
import pandas as pd
import matplotlib.pyplot as plt
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch


def count_sheets(excel_path):
    exc = xl.load_workbook(excel_path)
    return len(exc.sheetnames)


def sum_fields(excel_path):
    try:
        excel_data = pd.ExcelFile(excel_path)
        sum_total = 0

        for sheet_name in excel_data.sheet_names:
            df = pd.read_excel(excel_data, sheet_name)
            sum_total += df.sum().sum()  # Calculate the sum of all fields in the sheet

        return sum_total
    except Exception as e:
        return f"An error occurred: {str(e)}"


def sum_of_each_sheet(excel_path):
    try:
        excel_data = pd.ExcelFile(excel_path)
        sheet_sums = {}

        for sheet_name in excel_data.sheet_names:
            df = pd.read_excel(excel_data, sheet_name)
            sheet_sum = df.sum().sum()
            sheet_sums[sheet_name] = sheet_sum

        plt.bar(sheet_sums.keys(), sheet_sums.values())
        plt.xlabel('Sheet Name')
        plt.ylabel('Sum of Fields')
        plt.title('Sum of Each Sheet in Excel File')
        plt.xticks(rotation=45)
        plt.show()
    except Exception as e:
        print(f"An error occurred: {str(e)}")


def avg_sum_of_each_sheet(excel_path):
    try:
        excel_data = pd.ExcelFile(excel_path)
        sheet_sums = {}

        for sheet_name in excel_data.sheet_names:
            df = pd.read_excel(excel_data, sheet_name)
            if df is not None:

                int_columns = df.select_dtypes(include=['int']).columns
                int_df = df[int_columns]

                sheet_sum = int_df.values.sum() if not int_df.empty else 0
                sheet_avg = sheet_sum / int_df.size if int_df.size > 0 else 0

                sheet_sums['name'] = sheet_name
                sheet_sums["sum"] = int(sheet_sum)
                sheet_sums["avg"] = int(sheet_avg)
            else:
                print('df is none')
        return sheet_sums
    except Exception as e:
        print(f"An error occurred: {str(e)}")


def plot_average_of_each_sheet(excel_path):
    df = pd.read_excel(excel_path, sheet_name=None)  # Read all sheets of the Excel file

    average_values = []
    sheet_names = []

    for sheet_name, sheet_df in df.items():
        avg_value = sheet_df.mean().mean()  # Calculate the average of the sheet
        average_values.append(avg_value)
        sheet_names.append(sheet_name)

    plt.figure()
    plt.bar(sheet_names, average_values)
    plt.xlabel('Sheet Name')
    plt.ylabel('Average Value')
    plt.title('Average Value of Each Sheet')
    plt.xticks(rotation=45, ha="right")  # Rotate x-axis labels for better readability
    plt.show()


def generate_pdf_report(data, output_file):
    c = canvas.Canvas(output_file)

    c.drawString(100, 800, "PDF Report")

    y_position = 750
    for key, value in data.items():
        c.drawString(100, y_position, f"{key}: {value}")
        y_position -= 20

    c.save()


def add_image_to_pdf(image_path, output_file):
    c = canvas.Canvas(output_file, pagesize=A4)

    # Load the image and fit it into the PDF
    c.drawImage(image_path, inch, inch, width=400, height=400)
    c.save()


# Example image file path
image_path = "assets/avg.png"
# Output file name
output_file = "output.pdf"

# Add the image to the PDF
add_image_to_pdf(image_path, output_file)

# Example usage:
excel_file = 'new.xlsx'
sum_of_each_sheet(excel_file)
count_sheets(excel_file)
sum_fields(excel_file)
plot_average_of_each_sheet(excel_file)

# Example data for the report
data = {
    "part1": {
        1: "generate_report",
        2: "uploadfile",
        3: "create_pdf"
    },
    "part2": {
        "level A": {
            1: "count_sheets",
            2: "sum_fields",
            3: "sum_of_each_sheet",
            4: "plot_average_of_each_sheet"
        },
        "level B": "this"
    }
}

# Output file name
output_file = "end2.pdf"

# Generate the PDF report
generate_pdf_report(data, output_file)
