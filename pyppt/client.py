###############################################################################
# IPython/Javascript client for the remote notebook
#
# (c) Vladimir Filimonov, January 2018
###############################################################################
from __future__ import absolute_import
from builtins import str
import matplotlib.pyplot as plt
import json
import os
import base64
import uuid

try:
    from urllib import urlencode  # Python 2
except ImportError:
    from urllib.parse import urlencode  # Python 3

import pyppt.core as pyppt
from pyppt._ver_ import __version__, __author__, __email__, __url__


###############################################################################
# Javscript templates
###############################################################################
_html_div = '<div id="{id}" class="pyppt">[pyppt] Waiting for server response...</div>'

_js_init = """
<script>
function b64toBlob(b64Data, contentType, sliceSize) {
    // Based on https://stackoverflow.com/questions/16245767/
    contentType = contentType || '';
    sliceSize = sliceSize || 512;

    var byteCharacters = atob(b64Data);
    var byteArrays = [];

    for (var offset = 0; offset < byteCharacters.length; offset += sliceSize) {
        var slice = byteCharacters.slice(offset, offset + sliceSize);
        var byteNumbers = new Array(slice.length);
        for (var i = 0; i < slice.length; i++) {
            byteNumbers[i] = slice.charCodeAt(i);
        }
        var byteArray = new Uint8Array(byteNumbers);
        byteArrays.push(byteArray);
    }

    var blob = new Blob(byteArrays, {type: contentType});
    return blob;
};

function getResults(data, div_id) {
    // Return results to IPython kernel
    var kernel = Jupyter.notebook.kernel
    if (kernel) { kernel.execute('_results_pyppt_js_ = "' + data + '"'); }

    // Wait until DIV is created and then fill it
    var checkExist = setInterval(function() {
       if ($('#' + div_id).length) {
          document.getElementById(div_id).textContent = data;
          console.log("[pyppt] Server response: " + data);

          clearInterval(checkExist);
       }
    }, 100);
};
</script>
"""

_js_get = """$.get("{url}", function(data){{getResults(data, "{id}");}}); """

_js_post = """$.ajax({{
    url: "{url}",
    type: "POST",
    data: '{json}',
    contentType: "application/json; charset=utf-8",
    success: function(data){{getResults(data, "{id}");}},
}});
"""

_js_upload = """
var base64ImageContent = "{data}";
var blob = b64toBlob(base64ImageContent, 'image/png');
var formData = new FormData();
formData.append("picture", blob);

$.ajax({{
    url: "{url}",
    type: "POST",
    cache: false,
    contentType: false,
    processData: false,
    data: formData,
    success: function(data){{getResults(data, "{id}");}},
}});
"""

_js_upload_and_post = """
var base64ImageContent = "{data}";
var blob = b64toBlob(base64ImageContent, 'image/png');
var formData = new FormData();
formData.append("picture", blob);

$.ajax({{
    url: "{url1}",
    type: "POST",
    cache: false,
    contentType: false,
    processData: false,
    data: formData,
    success: function(data){{
        var new_data = JSON.parse('{json}');
        new_data["filename"] = data;

        $.ajax({{
            url: "{url2}",
            type: "POST",
            data: JSON.stringify(new_data),
            contentType: "application/json; charset=utf-8",
            success: function(data2){{getResults(data2, "{id}");}},
        }});
    }},
}});
"""


###############################################################################
# Client Classes
###############################################################################
class ClientGeneric(object):
    def __init__(self, host, port):
        self._url = 'http://%s:%s/' % (host, port)
        self._init_lib_()

    def _init_lib_(self):
        pass

    def url(self, method, **kwargs):
        res = self._url + method
        args = {_: kwargs[_] for _ in kwargs if kwargs[_] is not None}
        if len(args) > 0:
            res = res + '?' + urlencode(args)
        self._last_url = res
        return res

    def __getattr__(self, name):
        # to be called when Client is not initialized
        if name in ('get', 'post', 'upload_picture', 'post_and_figure'):
            raise Exception('Client was not initialized. Run init_client() first.')


_client = ClientGeneric('', '')


###############################################################################
class ClientJavascript(ClientGeneric):
    def _init_lib_(self):
        import IPython  # local references to library
        self.display = IPython.display

    def _div_id(self):
        """ Unique name for DIV """
        return 'pptdiv_%s' % (str(uuid.uuid4())[:8])

    def _run_js(self, script, **kwargs):
        idx = self._div_id()
        self._last_code = script.format(id=idx, **kwargs)

        # Trick posted to https://stackoverflow.com/questions/48248987/
        self.display.display(self.display.Javascript(self._last_code),
                             display_id=idx)
        self.display.display(self.display.HTML(_html_div.format(id=idx)),
                             display_id=idx, update=True)
        return None

    def init_js(self):
        self.display.display(self.display.HTML(_js_init))
        return None

    def get(self, method, **kwargs):
        return self._run_js(_js_get, url=self.url(method, **kwargs))

    def post(self, method, **kwargs):
        args = {_: kwargs[_] for _ in kwargs if kwargs[_] is not None}
        return self._run_js(_js_post, url=self.url(method), json=json.dumps(args))

    @staticmethod
    def _read_base64(filename, delete=False):
        with open(filename, 'rb') as f:
            data = f.read()
        data = base64.standard_b64encode(data)
        if delete:
            os.remove(filename)
        return str(data, 'utf-8')

    def upload_picture(self, filename, delete=False):
        return self._run_js(_js_upload, url=self.url('upload_picture'),
                            data=self._read_base64(filename, delete))

    def post_and_figure(self, method, filename, delete=True, **kwargs):
        """ Uploads figure to server and then call POST """
        args = {_: kwargs[_] for _ in kwargs if kwargs[_] is not None}
        return self._run_js(_js_upload_and_post, url1=self.url('upload_picture'),
                            url2=self.url(method), json=json.dumps(args),
                            data=self._read_base64(filename, delete))


###############################################################################
class ClientRequests(ClientGeneric):
    def _init_lib_(self):
        import requests  # local references to library
        self.requests = requests

    def get(self, method, **kwargs):
        r = self.requests.get(self.url(method, **kwargs))
        self._last_request = r
        return r.text

    def post(self, method, **kwargs):
        args = {_: kwargs[_] for _ in kwargs if kwargs[_] is not None}
        r = self.requests.post(self.url(method), json=args)
        self._last_request = r
        return r.text

    def upload_picture(self, filename, delete=False):
        with open(filename, 'rb') as f:
            r = self.requests.post(self.url('upload_picture'),
                                   files={'picture': f})
        self._last_request = r
        if delete:
            os.remove(filename)
        return r.text

    def post_and_figure(self, method, filename, delete=True, **kwargs):
        """ Uploads figure to server and then call POST """
        remote_fname = self.upload_picture(filename, delete=delete)
        return self.post(method, filename=remote_fname, **kwargs)


###############################################################################
def init_client(host=pyppt._LOCALHOST, port=pyppt._DEFAULT_PORT, javascript=True):
    """ Initialize client on the remote server.

        By default it will be using IPython notebook as a proxy and will embed
        javascripts in the notebook, that will be executed in browser on the
        local machine.

        If javascript is set to False, the client will try to connect to server
        running on the Windows machine directly. Then proper external IP address
        (or host name / url) and port should be specified, and firewalls on both
        client and server should be set.
    """
    global _client
    if javascript:
        _client = ClientJavascript(host, port)
    else:
        _client = ClientRequests(host, port)

    # Hijack matplotlib
    plt.add_figure = add_figure
    plt.replace_figure = replace_figure
    if javascript:
        return _client.init_js()


###############################################################################
# Exposed methods
###############################################################################
def title_to_front(slide_no=None):
    return _client.get('title_to_front', slide_no=slide_no)


def set_title(title, slide_no=None):
    return _client.get('set_title', title=title, slide_no=slide_no)


def set_subtitle(subtitle, slide_no=None):
    return _client.get('set_subtitle', subtitle=subtitle, slide_no=slide_no)


def add_slide(slide_no=None, layout_as=None, make_active=True):
    return _client.get('add_slide', slide_no=slide_no, layout_as=layout_as,
                       make_active=int(make_active))


def goto_slide(slide_no):
    return _client.get('goto_slide', slide_no=slide_no)


def get_shape_positions(slide_no=None):
    return _client.get('get_shape_positions', slide_no=slide_no)


def get_image_positions(slide_no=None):
    return _client.get('get_image_positions', slide_no=slide_no)


def get_slide_dimensions():
    return _client.get('get_slide_dimensions')


def get_notes():
    return _client.get('get_notes')


###############################################################################
def _save_figure(tight, **kwargs):
    # Save the figure to png in temporary directory
    fname = pyppt._temp_fname()
    if tight:
        # Usually is an overkill, but is needed sometimes...
        plt.tight_layout()
        plt.savefig(fname, bbox_inches='tight', **kwargs)
    else:
        plt.savefig(fname, **kwargs)
    w, h = plt.gcf().get_size_inches()
    return fname, w, h


def add_figure(bbox=None, slide_no=None, keep_aspect=True, tight=True,
               delete_placeholders=True, replace=False, **kwargs):
    fname, w, h = _save_figure(tight, **kwargs)
    return _client.post_and_figure('add_figure', filename=fname, bbox=bbox,
                                   slide_no=slide_no, keep_aspect=keep_aspect,
                                   delete_placeholders=delete_placeholders,
                                   replace=replace, w=w, h=h)


def replace_figure(pic_no=None, left_no=None, top_no=None, zorder_no=None,
                   slide_no=None, keep_zorder=True, tight=True, **kwargs):
    fname, w, h = _save_figure(tight, **kwargs)
    return _client.post_and_figure('replace_figure', filename=fname, pic_no=pic_no,
                                   left_no=left_no, top_no=top_no,
                                   zorder_no=zorder_no, slide_no=slide_no,
                                   keep_zorder=keep_zorder, w=w, h=h)


###############################################################################
title_to_front.__doc__ = pyppt.title_to_front.__doc__
set_title.__doc__ = pyppt.set_title.__doc__
set_subtitle.__doc__ = pyppt.set_subtitle.__doc__
add_slide.__doc__ = pyppt.add_slide.__doc__
goto_slide.__doc__ = pyppt.goto_slide.__doc__

get_shape_positions.__doc__ = pyppt.get_shape_positions.__doc__
get_image_positions.__doc__ = pyppt.get_image_positions.__doc__
get_slide_dimensions.__doc__ = pyppt.get_slide_dimensions.__doc__
get_notes.__doc__ = pyppt.get_notes.__doc__

add_figure.__doc__ = pyppt.add_figure.__doc__
replace_figure.__doc__ = pyppt.replace_figure.__doc__
