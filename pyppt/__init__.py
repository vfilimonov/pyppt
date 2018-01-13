# Python 2/3 compatibility
from __future__ import absolute_import

# Import library contents
from .core import *

# Metadata to be shared between module and setup.py
from ._ver_ import __version__, __author__, __email__, __url__

# Hijack matplotlib - methods for the local machine only
if win32client is not None:
    import matplotlib.pyplot as _plt
    _plt.add_figure = add_figure
    _plt.replace_figure = replace_figure
