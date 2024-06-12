<template>
  <div>
    <h1>Export Excel File with Chart Image</h1>
    <canvas id="chart" width="800" height="400"></canvas>
    <button @click="exportExcel">Export to Excel</button>
  </div>
</template>

<script>
import { saveAs } from 'file-saver';
import ExcelJS from 'exceljs';
import { Chart, BarController, BarElement, CategoryScale, LinearScale, Title, Tooltip, Legend } from 'chart.js';

Chart.register(BarController, BarElement, CategoryScale, LinearScale, Title, Tooltip, Legend);

export default {
  name: 'ExcelChart',
  data() {
    return {
      data: [
        { Date: '2024-01-01', 'Energy Produced (kWh)': 5.2, 'Peak Sunlight Hours': 4.5 },
        { Date: '2024-01-02', 'Energy Produced (kWh)': 5.0, 'Peak Sunlight Hours': 4.3 },
        { Date: '2024-01-03', 'Energy Produced (kWh)': 4.8, 'Peak Sunlight Hours': 4.2 },
        { Date: '2024-01-04', 'Energy Produced (kWh)': 5.5, 'Peak Sunlight Hours': 4.7 },
        { Date: '2024-01-05', 'Energy Produced (kWh)': 6.0, 'Peak Sunlight Hours': 5.0 }
      ]
    };
  },
  mounted() {
    this.renderChart();
  },
  methods: {
    renderChart() {
      const ctx = document.getElementById('chart').getContext('2d');
      new Chart(ctx, {
        type: 'bar',
        data: {
          labels: this.data.map(row => row.Date),
          datasets: [{
            label: 'Energy Produced (kWh)',
            data: this.data.map(row => row['Energy Produced (kWh)']),
            backgroundColor: 'rgba(75, 192, 192, 0.8)',
            borderColor: 'rgba(75, 192, 192, 1)',
            borderWidth: 1
          }]
        },
        options: {
          plugins: {
            title: {
              display: true,
              text: 'Energy Produced Over Time',
              font: {
                size: 24 // Increase the title font size
              }
            },
            legend: {
              labels: {
                font: {
                  size: 14 // Increase the legend font size
                }
              }
            },
            tooltip: {
              bodyFont: {
                size: 14 // Increase the tooltip body font size
              }
            }
          },
          scales: {
            x: {
              ticks: {
                font: {
                  size: 14 // Increase the x-axis ticks font size
                }
              }
            },
            y: {
              ticks: {
                font: {
                  size: 14 // Increase the y-axis ticks font size
                }
              }
            }
          }
        }
      });
    },
    exportExcel() {
      const data = this.data;

      // Create a workbook and a worksheet
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Solar Data');

      // Add column headers and rows
      worksheet.columns = [
        { header: 'Date', key: 'Date', width: 15 },
        { header: 'Energy Produced (kWh)', key: 'Energy Produced (kWh)', width: 25 },
        { header: 'Peak Sunlight Hours', key: 'Peak Sunlight Hours', width: 25 }
      ];
      data.forEach(row => {
        worksheet.addRow(row);
      });

      // Generate chart as image
      const canvas = document.getElementById('chart');
      canvas.toBlob(blob => {
        const reader = new FileReader();
        reader.onload = function () {
          const imageId = workbook.addImage({
            base64: reader.result,
            extension: 'png'
          });

          worksheet.addImage(imageId, {
            tl: { col: 6, row: 1 }, // Adjusted column to place the chart to the right of the data without using columns D and E
            ext: { width: 800, height: 400 }
          });

          // Save the workbook to a file
          workbook.xlsx.writeBuffer().then(buffer => {
            saveAs(new Blob([buffer]), 'SolarData.xlsx');
          });
        };
        reader.readAsDataURL(blob);
      }, 'image/png');
    }
  }
};
</script>

<style scoped>
button {
  padding: 10px 20px;
  font-size: 16px;
  background-color: #4CAF50;
  color: white;
  border: none;
  cursor: pointer;
  border-radius: 5px;
}

button:hover {
  background-color: #45a049;
}
</style>  