from flask import Flask , request , render_template_string , send_file
import pandas as pd
from openpyxl import Workbook , load_workbook 


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
    data = {
        "Name" : [],
        "Phone Number" : []

    }
    df = pd.DataFrame(data)
    file_path = "contacts.xlsx"
    df.to_excel(file_path, index=False)
    wb = load_workbook ('contacts.xlsx')
    ws=wb.active

    prefix = request.form.get("prefix")
    contact = request.form.get("contact")
    contact_numbers = contact.split(",")
    i = 1
    for num in contact_numbers : 
      new_row = {'A' : '{1}_{0}'.format(i , prefix) , 'b' : num}
      ws.append(new_row)
      wb.save('contacts.xlsx')
      i = i + 1

    return send_file("contacts.xlsx", as_attachment=True)



if __name__ == "__main__":
    app.run(debug=True)


