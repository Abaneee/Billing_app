import os
import json
from flask import Flask, render_template, request
from datetime import datetime
from num2words import num2words

try:
    import win32print
    import win32api
    USE_WINDOWS_PRINT = True
except ImportError:
    USE_WINDOWS_PRINT = False

app = Flask(__name__)
RST_FILE = "rst_no.json"

ESC = '\x1B'
BOLD_ON = ESC + 'E'
BOLD_OFF = ESC + 'F'

def load_rst_number():
    try:
        with open(RST_FILE, "r") as f:
            return int(json.load(f))
    except:
        return 1

def save_rst_number(rst_no):
    with open(RST_FILE, "w") as f:
        json.dump(rst_no, f)

@app.route('/')
def home():
    rst_no = load_rst_number()
    return render_template("form.html", rst_no=str(rst_no).zfill(4))

@app.route('/print', methods=['POST'])
def print_bill():
    rst_no = int(request.form["rst_no"])
    material = request.form["material"].upper()
    vehicle_no = request.form["vehicle_no"].upper()
    gross_wt = int(request.form["gross_wt"])
    tare_wt = int(request.form["tare_wt"])
    net_wt = gross_wt - tare_wt

    try:
        date_str = datetime.strptime(request.form["date"], "%d/%m/%Y").strftime("%d/%m/%Y")
    except:
        date_str = datetime.today().strftime("%d/%m/%Y")

    net_words = num2words(net_wt).replace(",", "").upper()
    divider = "-" * 80

    content = (
        f"{BOLD_ON}{'UDHAGAMANDALAM':^80}{BOLD_OFF}\n"
        f"{BOLD_ON}{'MUNICIPALITY':^80}{BOLD_OFF}\n"
        f"{BOLD_ON}{'THITUGUL COMPOST YARD':^80}{BOLD_OFF}\n"
        f"{BOLD_ON}{'OOTY':^80}{BOLD_OFF}\n\n"
        f"{'RST NO     : '}{BOLD_ON}{str(rst_no).zfill(4):<33}{BOLD_OFF}"
        f"{'VEHICLE NO : '}{BOLD_ON}{vehicle_no:<40}{BOLD_OFF}\n"
        f"{'MATERIAL   : ' + material:<80}\n"
        f"{divider}\n"
        f"{'GROSS WT   : '}{BOLD_ON}{str(gross_wt) + ' kg':<33}{BOLD_OFF}"
        f"{'DATE : ' + date_str:<40}\n"
        f"{'TARE  WT   : '}{BOLD_ON}{str(tare_wt) + ' kg':<33}{BOLD_OFF}"
        f"{'DATE : ' + date_str:<40}\n"
        f"{'NET   WT   : '}{BOLD_ON}{str(net_wt) + ' kg':<33}{BOLD_OFF}"
        f"{BOLD_ON}{net_words[:40]:<40}{BOLD_OFF}\n"
        f"{divider}\n"
        f"{'OPERATOR SIGNATURE:'}\n"
        f"{'-' * 80}\n\n"
        f"{'WELCOME':^80}\n"
    )

    file_path = "print_bill.txt"
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(content)

    if USE_WINDOWS_PRINT:
        printer = win32print.GetDefaultPrinter()
        win32api.ShellExecute(0, "printto", file_path, f'"{printer}"', ".", 0)
    else:
        os.system(f'notepad /p {file_path}')

    save_rst_number(rst_no + 1)

    return render_template("bill.html", content=content)

if __name__ == "__main__":
    app.run(debug=True)
