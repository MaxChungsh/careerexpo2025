from flask import Flask, render_template, request, send_file, redirect, url_for
import os
from processor import StudentAssignmentProcessor

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if files are uploaded
        response_file = request.files.get('response_file')
        classlist_file = request.files.get('classlist_file')

        if not response_file or not classlist_file:
            return "Please upload both files.", 400

        # Save uploaded files
        response_path = os.path.join(app.config['UPLOAD_FOLDER'], 'response_file.xlsx')
        classlist_path = os.path.join(app.config['UPLOAD_FOLDER'], 'classlist_file.xlsx')
        response_file.save(response_path)
        classlist_file.save(classlist_path)

        # Process files
        processor = StudentAssignmentProcessor('response_file.xlsx', 'classlist_file.xlsx')
        output_file = os.path.join(app.config['OUTPUT_FOLDER'], 'processed_output.xlsx')
        processor.process(output_file=output_file)

        return redirect(url_for('download'))

    return render_template('index.html')

@app.route('/download', methods=['GET'])
def download():
    output_file = os.path.join(app.config['OUTPUT_FOLDER'], 'processed_output.xlsx')
    if os.path.exists(output_file):
        return send_file(output_file, as_attachment=True)
    return "Output file not found.", 404

if __name__ == '__main__':
    app.run(debug=True)
