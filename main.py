import win32com.client as win32
import os
from helpers.helper_files import create_current_captiatlization_helper, \
                                 get_calendar_slide_helper, \
                                 create_ppt_with_template
from flask import Flask, send_file
import pythoncom
                                 
app = Flask(__name__)

if not os.path.exists(os.path.join(os.getcwd(), "output_path")):
    print("Creating an output path")
    os.makedirs("output_path", exist_ok=True)

def create_capitalization_slide(presentaion_template):
    presentaion_template = create_current_captiatlization_helper(presentaion_template)
    return presentaion_template

def create_calendar_slide(presentaion_template):
    presentaion_template = get_calendar_slide_helper(presentaion_template, date_str = '2025-06-01')
    return presentaion_template

@app.route('/capitalization-slide')
def download_ppt_capitalization():
    pythoncom.CoInitialize()
    template_path = os.path.join(os.getcwd(), r"template_path\SampleTemplate.pptx")
    presentation = create_ppt_with_template(template_path)
    output_path = create_capitalization_slide(presentation)
    pythoncom.CoUninitialize()
    filename = output_path.split('\\')[-1]
    return send_file(output_path, as_attachment=True, download_name=filename)

@app.route('/calendar-slide')
def download_ppt_calendar():
    pythoncom.CoInitialize()
    template_path = os.path.join(os.getcwd(), r"template_path\SampleTemplate.pptx")
    presentation = create_ppt_with_template(template_path)
    output_path = create_calendar_slide(presentation)
    pythoncom.CoUninitialize()
    filename = output_path.split('\\')[-1]
    return send_file(output_path, as_attachment=True, download_name=filename)

if __name__ == "__main__":

    app.run(debug=True)
    ## http://localhost:5000/capitalization-slide
    ## http://localhost:5000/calendar-slide
    #presentation = create_calendar_slide(presentation)
