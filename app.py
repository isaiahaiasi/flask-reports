import os
from flask import Flask, json, request, flash, redirect, url_for, send_from_directory
from werkzeug.utils import secure_filename


ALLOWED_EXTENSIONS = {'csv'}
UPLOAD_FOLDER = './www'

# create www folder
os.mkdir('./www')


app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


def is_allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# receives the raw timesheets csv
# returns the formatted timesheets xlsx
@app.route("/format-timesheet", methods=['POST', 'GET'])
def format_timesheet():
    if request.method == 'POST':
        if 'file' not in request.files:
            # TODO: replace with error response
            flash('No file part')
            return redirect(request.url)

        f = request.files['file']

        # if user doesn't select a file, browser submits empty unamed file
        if f.filename == '':
            # TODO: replace with error response
            flash('No selected file')
            return redirect(request.url)

        if f and is_allowed_file(f.filename):
            filename = secure_filename(f.filename)
            f.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            return redirect(url_for('download_file', name=filename))

    return '''
    <!doctype html>
    <title>Upload new file</title>
    <h1>Upload new File</h1>
    <form method=post enctype=multipart/form-data>
        <input type=file name=file>
        <input type=submit value=Upload>
    </form>
    '''


@app.route('/uploads/<name>')
def download_file(name):
    return send_from_directory(app.config['UPLOAD_FOLDER'], name)


# base endpoint, can be used to wake up a heroku dyno
@app.route("/")
def hello_world():
    return json.jsonify(message="wakeup complete")
