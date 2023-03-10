from flask import Flask, render_template, request, redirect, url_for
import pdfplumber
import re
import pandas as pd

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if a file was uploaded
        if 'file' not in request.files:
            return redirect(request.url)

        file = request.files['file']

        # Check if the file is a PDF
        if file.filename.endswith('.pdf'):
            # Read the PDF file
            with pdfplumber.open(file) as pdf:
                data = []  # create an empty list to store the extracted data

                for page in pdf.pages:
                    page_data = {}  # create an empty dictionary to store the extracted data from the page
                    text = page.extract_text()

                    # Remove Barcode label and last line
                    text = text.split("\n")[:-2]
                    text = "\n".join(text)
                    text = re.sub(r'Barcode:', '', text)

                    # Select lines
                    line_start = re.compile(r'^\w+(?:\s+\w+)?:(.*)$')
                    for line in text.split("\n"):
                        if line_start.match(line):
                            # Extract the key and value from the line
                            key, value = line.split(":", 1)
                            key = key.strip()
                            value = value.strip()

                            # Convert the date strings to datetime objects
                            if key == "Date":
                                value = pd.to_datetime(value)

                            # Add the key-value pair to the page_data dictionary
                            page_data[key] = value

                    # Add the page_data dictionary to the data list
                    data.append(page_data)

                # Convert the list of dictionaries to a pandas DataFrame
                df = pd.DataFrame(data)

                # Drop the "Description" column
                df = df.drop(columns=["Description"])

                # Save the DataFrame to an Excel file
                df.to_excel("lending_requests.xlsx", index=False)

            return redirect(url_for('index'))

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
