# Dummy __init__ file
from .pyppt import *

# Hijack matplotlib
import matplotlib.pyplot as _plt
_plt.add_figure = add_figure
