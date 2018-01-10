###############################################################################
# IPython/Javascript client for the remote notebook
#
# (c) Vladimir Filimonov, January 2018
###############################################################################
import matplotlib.pyplot as plt
import json

try:
    import IPython
except ImportError:
    IPython = None
try:
    import requests
except ImportError:
    requests = None

import pyppt as pyppt
from ._ver_ import __version__, __author__, __email__, __url__


###############################################################################
class ClientJavascript(object):
    def __init__(self, host, port):
        import IPython  # To throw ImportError if absent
        self.url = 'http://%s:%s/' % (host, port)

    def get(self, method, **kwargs):
        raise NotImplementedError  # TODO

    def post(self, method, **kwargs):
        raise NotImplementedError  # TODO

    def upload_picture(self, filename, delete=False):
        raise NotImplementedError  # TODO

    def post_and_figure(self, method, filename, delete=True, **kwargs):
        """ Uploads figure to server and then call POST """
        raise NotImplementedError  # TODO


###############################################################################
class ClientRequests(object):
    def __init__(self, host, port):
        import requests  # To throw ImportError if absent
        self.url = 'http://%s:%s' % (host, port)
        raise NotImplementedError  # TODO

    def get(self, method, **kwargs):
        raise NotImplementedError  # TODO

    def post(self, method, **kwargs):
        raise NotImplementedError  # TODO

    def upload_picture(self, filename, delete=False):
        raise NotImplementedError  # TODO

    def post_and_figure(self, method, filename, delete=True, **kwargs):
        """ Uploads figure to server and then call POST """
        raise NotImplementedError  # TODO


###############################################################################
def init_client(host='127.0.0.1', port='5000', javascript=True):
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


###############################################################################
# Exposed methods
###############################################################################
def title_to_front(slide_no=None):
    """ Bring title and subtitle to front """
    return _client.get('title_to_front', slide_no=slide_no)


def set_title(title, slide_no=None):
    """ Set title for the slide (active or of a given number).
        If slide contain multiple Placeholder/Title objects, only first one is set.
    """
    return _client.get('set_title', title=title, slide_no=slide_no)


def set_subtitle(subtitle, slide_no=None):
    """ Set title for the slide (active or of a given number).
        If slide contain multiple Placeholder/Title objects, only first one is set.
    """
    return _client.get('set_subtitle', subtitle=subtitle, slide_no=slide_no)


def add_slide(slide_no=None, layout_as=None):
    """ Add slide after slide number "slide_no" with the layout as in the slide
        number "layout_as".
        If "slide_no" is None, new slide will be added after the active one.
        If "layout_as" is None, new slide will have layout as the active one.
        Returns the number of the added slide.
    """
    return _client.get('add_slide', slide_no=slide_no, layout_as=layout_as)


###############################################################################
def get_shape_positions(slide_no=None):
    """ Get positions of all shapes in the slide.
        Return list of lists of the format [x, y, w, h, type].
    """
    return _client.get('get_shape_positions', slide_no=slide_no)


def get_image_positions(slide_no=None):
    """ Get positions of all images in the slide.
        Return list of lists of the format [x, y, w, h].
    """
    return _client.get('get_image_positions', slide_no=slide_no)


def get_slide_dimensions():
    """ Get width and heights of the slide """
    return _client.get('get_slide_dimensions')


def get_notes():
    """ Extract notes for all slides from the presentation """
    return _client.get('get_notes')


###############################################################################
###############################################################################
def add_figure(bbox=None, slide_no=None, keep_aspect=True, tight=True,
               delete_placeholders=True, replace=False, **kwargs):
    # Save the figure to png in temporary directory
    fname = pyppt._temp_fname()
    if tight:
        # Usually is an overkill, but is needed sometimes...
        plt.tight_layout()
        plt.savefig(fname, bbox_inches='tight', **kwargs)
    else:
        plt.savefig(fname, **kwargs)

    return _client.post_and_figure('add_figure', filename=fname, bbox=bbox,
                                   slide_no=slide_no, keep_aspect=keep_aspect,
                                   delete_placeholders=delete_placeholders,
                                   replace=replace, **kwargs)


###############################################################################
def replace_figure(pic_no=None, left_no=None, top_no=None, zorder_no=None,
                   slide_no=None, keep_zorder=True, tight=True, **kwargs):
    # Save the figure to png in temporary directory
    fname = pyppt._temp_fname()
    if tight:
        # Usually is an overkill, but is needed sometimes...
        plt.tight_layout()
        plt.savefig(fname, bbox_inches='tight', **kwargs)
    else:
        plt.savefig(fname, **kwargs)

    return _client.post_and_figure('replace_figure', filename=fname, pic_no=pic_no,
                                   left_no=left_no, top_no=top_no,
                                   zorder_no=zorder_no, slide_no=slide_no,
                                   keep_zorder=keep_zorder, **kwargs)


###############################################################################
add_figure.__doc__ = pyppt.add_figure.__doc__
replace_figure.__doc__ = pyppt.replace_figure.__doc__
