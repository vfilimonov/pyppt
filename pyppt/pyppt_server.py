###############################################################################
# Flask server for the local Windows machine
#
# (c) Vladimir Filimonov, January 2018
###############################################################################
from __future__ import print_function
from flask import Flask, request, jsonify
from flask_cors import CORS
import json

import pyppt as pyppt


###############################################################################
pyppt._check_win32com()

app = Flask(__name__)
CORS(app)


###############################################################################
@app.route('/')
def home():
    from _ver_ import __author__, __email__, __url__
    return ('<html><body>pyppt connector by %s (<a href="mailto:%s">%s</a>)'
            '<br>See docs at <a href="%s">%s</a></body></html>'
            % (__author__, __email__, __email__, __url__, __url__))


###############################################################################
# Exposed via GET
###############################################################################
@app.route('/title_to_front')
def title_to_front():
    print('title_to_front', request.args.to_dict())
    slide_no = request.args.get('slide_no', default=None, type=int)
    pyppt.title_to_front(slide_no=slide_no)
    return 'OK'


@app.route('/set_title')
def set_title():
    print('set_title', request.args.to_dict())
    title = request.args.get('title', default='', type=str)
    slide_no = request.args.get('slide_no', default=None, type=int)
    pyppt.set_title(title=title, slide_no=slide_no)
    return 'OK'


@app.route('/set_subtitle')
def set_subtitle():
    print('set_subtitle', request.args.to_dict())
    subtitle = request.args.get('subtitle', default='', type=str)
    slide_no = request.args.get('slide_no', default=None, type=int)
    pyppt.set_subtitle(subtitle=subtitle, slide_no=slide_no)
    return 'OK'


@app.route('/add_slide')
def add_slide():
    print('add_slide', request.args.to_dict())
    slide_no = request.args.get('slide_no', default=None, type=int)
    layout_as = request.args.get('layout_as', default=None, type=int)
    pyppt.add_slide(layout_as=layout_as, slide_no=slide_no)
    return 'OK'


###############################################################################
@app.route('/get_shape_positions')
def get_shape_positions():
    print('get_shape_positions', request.args.to_dict())
    slide_no = request.args.get('slide_no', default=None, type=int)
    return repr(pyppt.get_shape_positions(slide_no=slide_no))


@app.route('/get_image_positions')
def get_image_positions():
    print('get_image_positions', request.args.to_dict())
    slide_no = request.args.get('slide_no', default=None, type=int)
    return repr(pyppt.get_image_positions(slide_no=slide_no,
                                          asarray=True, decimals=1).to_list())


@app.route('/get_slide_dimensions')
def get_slide_dimensions():
    return repr(pyppt.get_slide_dimensions())


@app.route('/get_notes')
def get_notes():
    return repr(pyppt.get_notes())


###############################################################################
# Exposed via POST
###############################################################################
@app.route('/upload_picture', methods=['POST'])
def upload_picture():
    raise NotImplementedError  # TODO
    if 'picture' not in request.files:
        raise Exception('No file part in the POST')
    filedata = request.files['picture']
    if filedata.filename == '':
        raise Exception('No selected file in the POST')

    fname = pyppt._temp_fname()
    filedata.save(fname)
    return fname


@app.route('/add_figure', methods=['POST'])
def add_figure():
    if request.is_json():
        args = request.get_json()
    else:
        raise Exception('Arguments are expected to be in JSON format')
    print('add_figure', args)
    # Parse JSON arguments
    fname = args['filename']
    bbox = args.get('bbox', None)
    slide_no = args.get('slide_no', None)
    keep_aspect = args.get('keep_aspect', True)
    replace = args.get('replace', False)
    delete_placeholders = args.get('delete_placeholders', True)
    target_z_order = args.get('target_z_order', None)
    # TODO: call add_figure
    return 'OK'


@app.route('/replace_figure', methods=['POST'])
def replace_figure():
    if request.is_json():
        args = request.get_json()
    else:
        raise Exception('Arguments are expected to be in JSON format')
    print('replace_figure', args)
    # Parse JSON arguments
    fname = args['filename']
    pic_no = args.get('pic_no', None)
    left_no = args.get('left_no', None)
    top_no = args.get('top_no', None)
    zorder_no = args.get('zorder_no', None)
    slide_no = args.get('slide_no', None)
    keep_aspect = args.get('keep_aspect', True)
    keep_zorder = args.get('keep_zorder', True)
    # TODO: call replace_figure
    return 'OK'


###############################################################################
# Run the server
# TODO: command-line arguments: http://flask.pocoo.org/snippets/133/
app.run()
