###############################################################################
# Flask server for the local Windows machine
#
# (c) Vladimir Filimonov, January 2018
###############################################################################
from __future__ import print_function, absolute_import
from flask import Flask, request, jsonify
from flask_cors import CORS
import json
import optparse

import pyppt.core as pyppt


###############################################################################
app = Flask(__name__)
CORS(app)


###############################################################################
@app.route('/')
def home():
    try:
        pyppt._check_win32com()
        ok = ''
    except Exception as e:
        ok = '<br><br><font color="#ff0033">NB! ' + str(e) + '</font>'
    msg = ('<html><body>pyppt connector ver. %s by %s (<a href="mailto:%s">%s</a>)'
           '<br>See docs at <a href="%s">%s</a>%s</body></html>'
           % (pyppt.__version__, pyppt.__author__, pyppt.__email__,
              pyppt.__email__, pyppt.__url__, pyppt.__url__, ok))
    return msg


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
    make_active = bool(request.args.get('make_active', default=None, type=int))
    pyppt.add_slide(layout_as=layout_as, slide_no=slide_no, make_active=make_active)
    return 'OK'


@app.route('/goto_slide')
def goto_slide():
    print('goto_slide', request.args.to_dict())
    slide_no = request.args.get('slide_no', default=None, type=int)
    pyppt.goto_slide(slide_no=slide_no)


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
    return repr(pyppt.get_image_positions(slide_no=slide_no))


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
    if request.is_json:
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
    w = args.get('w', None)
    h = args.get('h', None)
    # Call pyppt method
    pyppt._add_figure(fname, bbox=bbox, slide_no=slide_no, replace=replace,
                      keep_aspect=keep_aspect, target_z_order=target_z_order,
                      delete_placeholders=delete_placeholders,
                      delete=True, w=w, h=h)
    return 'OK'


@app.route('/replace_figure', methods=['POST'])
def replace_figure():
    if request.is_json:
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
    w = args.get('w', None)
    h = args.get('h', None)
    # Call pyppt method
    pyppt._replace_figure(fname, pic_no=pic_no, left_no=left_no, top_no=top_no,
                          zorder_no=zorder_no, slide_no=slide_no,
                          keep_zorder=keep_zorder, keep_aspect=keep_aspect,
                          delete=True, w=w, h=h)
    return 'OK'


###############################################################################
def flaskrun(app, default_host=pyppt._LOCALHOST, default_port=pyppt._DEFAULT_PORT):
    """ Runs Flask instance using command line arguments """
    # Based on http://flask.pocoo.org/snippets/133/
    parser = optparse.OptionParser()
    parser.add_option("-H", "--host",
                      help="Hostname of the Flask app [default %s]" % default_host,
                      default=default_host)
    parser.add_option("-P", "--port",
                      help="Port for the Flask app [default %s]" % default_port,
                      default=default_port)
    parser.add_option("-d", "--debug",
                      action="store_true", dest="debug",
                      help=optparse.SUPPRESS_HELP)

    options, _ = parser.parse_args()
    app.run(debug=options.debug, host=options.host, port=int(options.port))


def pyppt_server():
    flaskrun(app)


###############################################################################
if __name__ == '__main__':
    pyppt_server()
