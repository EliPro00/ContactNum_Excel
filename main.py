from flask import Flask, request, render_template_string, send_file
import pandas as pd
from openpyxl import Workbook
from io import BytesIO

app = Flask(__name__)

HTML_Template = """
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <link
      rel="stylesheet"
      href="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css" />
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.7.1/jquery.min.js"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/js/bootstrap.min.js"></script>
    <title>Document</title>
  </head>
 <body>
    <div class="container">
        <h2>Generate and Download Excel File</h2>
        <form action="/capture_data" method="POST">
            <div class="form-group">
                <label for="prefix">Contact Pre-Fix</label>
                <input type="text" class="form-control" id="prefix" name="prefix" placeholder="Enter prefix" required>
            </div>
            <div class="form-group">
                <label for="contact">Contact Numbers (comma-separated)</label>
                <input type="text" class="form-control" id="contact" name="contact" placeholder="Enter contact numbers" required>
            </div>
            <button type="submit" class="btn btn-primary">Generate Excel</button>
        </form>
    </div>
</body>
</html>
"""

@app.route("/")
def index():
    return render_template_string(HTML_Template)

@app.route('/capture_data', methods=['POST'])
def capture_data():
    # Get input from form
    prefix = request.form.get("prefix", "").strip()
    contact = request.form.get("contact", "").strip()

    # Validate input
    if not prefix or not contact:
        return "Prefix and Contact numbers are required!", 400

    # Split contacts
    contact_numbers = [num.strip() for num in contact.split(",") if num.strip()]

    if not contact_numbers:
        return "Invalid contact numbers!", 400

    # Create an Excel file in memory
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.append(["Name", "Phone Number"])  # Header row

    # Append data to Excel file
    for i, num in enumerate(contact_numbers, start=1):
        ws.append([f"{prefix}_{i}", num])

    # Save to in-memory buffer
    wb.save(output)
    output.seek(0)

    # Send file to user
    return send_file(output, as_attachment=True, download_name="contacts.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(debug=True)
