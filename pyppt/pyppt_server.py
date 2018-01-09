###############################################################################
# Flask server for the local Windows machine
#
# (c) Vladimir Filimonov, January 2018
###############################################################################
from flask import Flask, request, jsonify
from flask_cors import CORS
import json

import pyppt as pyppt


pyppt._check_win32com()

app = Flask(__name__)
CORS(app)


###############################################################################
@app.route('/')
def home():
    from pyppt import __author__, __email__, __url__
    return ('<html><body>pyppt connector by %s (<a href="mailto:%s">%s</a>)'
            '<br>See docs at <a href="%s">%s</a></body></html>'
            % (__author__, __email__, __email__, __url__, __url__))


###############################################################################
# Exposed via GET
###############################################################################
@app.route('/title_to_front')
def title_to_front():
    pyppt.title_to_front(**request.args.to_dict())
    return 'OK'


@app.route('/set_title')
def set_title():
    pyppt.set_title(**request.args.to_dict())
    return 'OK'


@app.route('/set_subtitle')
def set_subtitle():
    pyppt.set_subtitle(**request.args.to_dict())
    return 'OK'


@app.route('/add_slide')
def add_slide():
    pyppt.add_slide(**request.args.to_dict())
    return 'OK'


###############################################################################
@app.route('/get_shape_positions')
def get_shape_positions():
    return pyppt.get_shape_positions(**request.args.to_dict())


@app.route('/get_image_positions')
def get_image_positions():
    return pyppt.get_image_positions(**request.args.to_dict())


@app.route('/get_slide_dimensions')
def get_slide_dimensions():
    return pyppt.get_slide_dimensions()


@app.route('/get_notes')
def get_notes():
    return pyppt.get_notes()


###############################################################################
# Exposed via POST
###############################################################################
@app.route('/post', methods=['POST'])
def post():
    if request.is_json():
        args = request.get_json()
    else:
        raise Exception('Arguments are expected to be in JSON format')
    print(args)
    raise NotImplementedError  # TODO
    return 'OK'


@app.route('/upload_picture', methods=['POST'])
def upload_picture():
    raise NotImplementedError  # TODO


@app.route('/add_figure', methods=['POST'])
def add_figure():
    raise NotImplementedError  # TODO


@app.route('/replace_figure', methods=['POST'])
def replace_figure():
    raise NotImplementedError  # TODO


###############################################################################
# Run the server
# TODO: command-line arguments: http://flask.pocoo.org/snippets/133/
app.run()
