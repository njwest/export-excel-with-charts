# Export Excel with Chart Image Examples

## Client-side: Vue with ExcelJS and ChartJS

In `/vue_example` you will find a quick-and-dirty VueJS app that exports an Excel sheet containing a **static** ChartJS-generated image taken from the DOM.

See it in live: https://vue-excel-chart-output-demo.netlify.app/

**Note:** This is a quick and dirty demonstration, the ChartJS chart has responsive-width bars by default and doesn't have much of a container
so it will be exported as you see it vis-a-vis your current device width.

## Server-side: Python

In `/py_example` you will find a simple Python script that generates an Excel sheet 
containing a **static** chart image.

This Python script could be modified to send a file blob to a client rather than write a file to disk.

This Python script could be run and supervised by an Elixir server using [ErlPort](https://github.com/erlport/erlport).

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

4. Open the Excel sheet exported to the 
