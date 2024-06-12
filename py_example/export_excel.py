import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
import openpyxl

def create_excel_with_chart():
    # Dummy data for the solar array
    data = {
        'Date': ['2024-01-01', '2024-01-02', '2024-01-03', '2024-01-04', '2024-01-05'],
        'Energy Produced (kWh)': [5.2, 5.0, 4.8, 5.5, 6.0],
        'Peak Sunlight Hours': [4.5, 4.3, 4.2, 4.7, 5.0]
    }
    df = pd.DataFrame(data)

    # Create an Excel writer
    excel_buffer = BytesIO()
    with pd.ExcelWriter('SolarData.xlsx', engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Solar Data')
        workbook = writer.book
        worksheet = writer.sheets['Solar Data']

        # Generate a chart
        fig, ax = plt.subplots()
        df.plot(x='Date', y='Energy Produced (kWh)', kind='bar', ax=ax)
        chart_buffer = BytesIO()
        plt.savefig(chart_buffer, format='png')
        plt.close(fig)

        # Insert the chart into the Excel file
        img = openpyxl.drawing.image.Image(chart_buffer)
        img.anchor = 'E1'
        worksheet.add_image(img)

if __name__ == "__main__":
    create_excel_with_chart()
    print("Excel file with chart has been created successfully!")
