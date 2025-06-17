import os
from capitalization_svg import add_Captialization_svg
from calendar_svg import create_calendar_svg
from flask import Flask, send_file, jsonify
                                 
app = Flask(__name__)



if not os.path.exists(os.path.join(os.getcwd(), "output_path")):
    print("Creating an output path")
    os.makedirs("output_path", exist_ok=True)



@app.route('/capitalization-svg-download')
def download_svg_capitalization():
    filepath = "captalization_table.svg"
    output_path = os.path.join(os.path.join(os.getcwd(), "output_path"), filepath)
    add_Captialization_svg(filename = output_path)
    return send_file(output_path, as_attachment=True, download_name=filepath)



@app.route('/calendar-svg-download')
def download_svg_calendar():
    filepath = "calendar_timeline.svg"
    output_path = os.path.join(os.path.join(os.getcwd(), "output_path"), filepath)
    create_calendar_svg(filename = output_path)
    return send_file(output_path, as_attachment=True, download_name=filepath)



@app.route('/calendar-svg-source')
def calendar_svg_source():
    filepath = "calendar_timeline.svg"
    output_path = os.path.join(os.path.join(os.getcwd(), "output_path"), filepath)
    create_calendar_svg(filename = output_path)
    with open(output_path, "r") as fo:
        source_code = fo.read()
    
    response_obj = jsonify({
        "source_code": source_code
    })
    return source_code



@app.route('/capitalization-svg-source')
def capitalization_svg_source():
    filepath = "captalization_table.svg"
    output_path = os.path.join(os.path.join(os.getcwd(), "output_path"), filepath)
    add_Captialization_svg(filename = output_path)
    with open(output_path, "r") as fo:
        source_code = fo.read()
    
    return jsonify({
        "source_code": source_code
    })


if __name__ == "__main__":
    app.run(debug=True)
    
