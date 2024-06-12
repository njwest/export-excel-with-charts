# Export Excel with Chart Image Examples

This repo contains two quick examples of how to export Excel files containing charts,
one that works on the client-side with **VueJS** and one that works server-side from a **Python** script.

## Client-side: Vue with ExcelJS and ChartJS

In `/vue_example` you will find a quick-and-dirty VueJS app that exports an Excel sheet containing a **static** ChartJS-generated image taken from the DOM.

Demo: https://vue-excel-chart-output-demo.netlify.app/

**Note:** This is a quick and dirty demonstration, the ChartJS chart has responsive-width bars by default and doesn't have much of a container
so its text/bars will be exported relative to how it is displayed on the DOM -- meaning the text may be quite small.

## Server-side: Python

In `/py_example` you will find a simple Python script that generates an Excel sheet 
containing a **static** chart image.

This Python script could be modified to send a file blob to a client rather than write a file to disk.

This Python script could also be run and supervised by an Elixir server using [ErlPort](https://github.com/erlport/erlport).

## Running the Script

1. Create a virtual env and activate it:

```bash
cd ./py_example
python3 -m venv venv
source venv/bin/activate
```

2. Install dependencies:

```bash
pip install -r ./requirements.txt
```

3. Run the script:

```bash
python3 export_excel.py
```

This script creates an Excel sheet in this directory.

4. Open the Excel sheet and enjoy the chart!
