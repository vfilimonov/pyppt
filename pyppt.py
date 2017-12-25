###############################################################################
# Python module to insert figures to the open PPT
#
# (c) Vladimir Filimonov, December 2017
###############################################################################
from win32com import client
import matplotlib.pyplot as plt
import numpy as np
import warnings
import tempfile

__version__ = '0.1'
__author__ = 'Vladimir Filimonov'
__email__ = 'vladimir.a.filimonov@gmail.com'

###############################################################################
# Some constants from MSDN MS Office reference
# See also VBA reference: https://github.com/OfficeDev/VBA-content
###############################################################################
# MsoShapeType: https://msdn.microsoft.com/en-us/library/aa432678.aspx
msoPicture = 13
msoPlaceholder = 14
# PpPlaceholderType: https://msdn.microsoft.com/en-us/VBA/PowerPoint-VBA/articles/ppplaceholdertype-enumeration-powerpoint
ppPlaceholderTitle = 1
ppPlaceholderBody = 2
ppPlaceholderSubtitle = 4
# MsoZOrderCmd: https://msdn.microsoft.com/en-us/library/aa432726.aspx
msoBringToFront = 0
# Temporary text to be filled in empty placeholders
TEMPTEXT = '--TO-BE-REMOVED--'


###############################################################################
def _get_application():
    """ Get reference to PowerPoint application """
    Application = client.Dispatch('PowerPoint.Application')
    # Make it visible
    Application.Visible = True
    return Application


def get_active_slide():
    """ Get reference to active slide """
    return _get_application().ActiveWindow.View.Slide


def get_active_presentation():
    """ Get reference to active presentation """
    return _get_application().ActivePresentation


###############################################################################
def fill_empty_placeholders(Slide=None, text='Temp'):
    """ Dirty hack: fill all empty placeholders with some text.
        Returns a list of objects that were filled, so then the text could be
        cleared from them (see empty_placeholders()).
        If we don't do this - when we insert a figure, it will take place of
        the placeholder, rather than using specified coordinates.
    """
    if Slide is None:
        Slide = get_active_slide()
    filled = []
    for ii in range(Slide.Shapes.Count):
        item = Slide.Shapes.Item(1+ii)
        if item.Type == msoPlaceholder:
            ptype = item.PlaceholderFormat.type
            if ptype not in (ppPlaceholderTitle, ppPlaceholderSubtitle,
                             ppPlaceholderBody):
                if item.TextFrame.TextRange.Length == 0:  # if empty
                    item.TextFrame.TextRange.Text = TEMPTEXT
                    filled.append(item)
    return filled


def empty_filled_placeholders(items):
    """ Remove text from all placeholders that were filled by
        fill_empty_placeholders()
    """
    for item in items:
        item.TextFrame.TextRange.Text = ''


def delete_empty_placeholders(Slide=None):
    """ Delete all empty placeholders except Title and Subtitle """
    if Slide is None:
        Slide = get_active_slide()
    # we're going ro remove => iterate in reverse order
    for ii in range(Slide.Shapes.Count)[::-1]:
        item = Slide.Shapes.Item(1+ii)
        if item.Type == msoPlaceholder:
            ptype = item.PlaceholderFormat.type
            if ptype not in (ppPlaceholderTitle, ppPlaceholderSubtitle,
                             ppPlaceholderBody):
                if item.TextFrame.TextRange.Length == 0:  # if empty
                    item.delete()


def title_to_front(Slide=None):
    """ Bring title and subtitle to front """
    if Slide is None:
        Slide = get_active_slide()
    for ii in range(Slide.Shapes.Count):
        item = Slide.Shapes.Item(1+ii)
        if item.Type == msoPlaceholder:
            ptype = item.PlaceholderFormat.type
            if ptype not in (ppPlaceholderTitle, ppPlaceholderSubtitle,
                             ppPlaceholderBody):
                item.ZOrder(msoBringToFront)


###############################################################################
def get_shape_positions(Slide=None):
    """ Get positions of all shapes in the Slide.
        Return list of lists of the format [x, y, w, h, type].
    """
    if Slide is None:
        Slide = get_active_slide()
    positions = []
    for ii in range(Slide.Shapes.Count):
        item = Slide.Shapes.Item(1+ii)
        positions.append([item.Left, item.Top, item.Width, item.Height, item.Type])
    return positions


def get_image_positions(Slide=None, asarray=True, decimals=1):
    """ Get positions of all images in the Slide. If necessary, rounds
        coordinates to a given decimals (if "decimals" is not None)
        Return list of lists of the format [x, y, w, h].
    """
    positions = get_shape_positions(Slide)
    # Kepp only images
    positions = [p[:-1] for p in positions if p[-1] == msoPicture]
    if asarray:
        if decimals is not None:
            return np.round(np.array(positions), decimals=decimals)
        else:
            return np.array(positions)
    else:
        return positions


def get_slide_dimensions(Presentation=None):
    """ Get width and heights of the slide """
    if Presentation is None:
        Presentation = get_active_presentation()
    return (Presentation.PageSetup.SlideWidth,
            Presentation.PageSetup.SlideHeight)


def get_notes(Presentation=None):
    """ Extract notes for all slides from the presentation """
    if Presentation is None:
        Presentation = get_active_presentation()
    Slides = Presentation.Slides
    notes = []
    for ii in range(len(Slides)):
        notes.append(Slides[ii].NotesPage.Shapes.Placeholders[2]
                               .TextFrame.TextRange.Text)
    return notes


###############################################################################
def parse_bbox(bbox):
    """ Human-readable bbox-dimensions"""
    # Stub
    return bbox


###############################################################################
def _temp_fname():
    """ Return a name of a temporary file """
    f = tempfile.NamedTemporaryFile(delete=False)
    f.close()
    return f.name + '.png'


###############################################################################
def add_figure(bbox, keep_aspect=True, delete_placeholders=True,
               bbox_inches='tight', **kwargs):
    """ Add current figure to the active slide """
    # Save the figure to png
    fname = _temp_fname()
    if bbox_inches == 'tight':
        # Usually is an overkill, but is needed sometimes...
        plt.tight_layout()
    plt.savefig(fname, bbox_inches=bbox_inches, **kwargs)

    # Parse bbox name if necessary
    bbox = parse_bbox(bbox)

    # If keep_aspect:
    if keep_aspect:
        w, h = np.asfarray(plt.gcf().get_size_inches())
        bx, by, bw, bh = np.asfarray(bbox)
        aspect_fig = w / h
        aspect_bbox = bw / bh
        if aspect_fig > aspect_bbox:  # Figure is wider than bbox
            newh = bw / aspect_fig
            bbox = [bx, by + bh/2. - newh/2., bw, newh]
        else:
            neww = bh * aspect_fig
            bbox = [bx + bw/2. - neww/2., by, neww, bh]

    # Now insert to PowerPoint
    Slide = get_active_slide()
    if delete_placeholders:
        delete_empty_placeholders(Slide)
    else:
        items = fill_empty_placeholders(Slide)
    shape = Slide.Shapes.AddPicture(FileName=fname, LinkToFile=False,
                                    SaveWithDocument=True, Left=bbox[0],
                                    Top=bbox[1], Width=bbox[2], Height=bbox[3])
    if not (shape.Left == bbox[0] and shape.Top == bbox[1] and
            shape.Width == bbox[2] and shape.Height == bbox[3]):
        warnings.warn('BBox of the inserted figure was not respected: '
                      '(%.1f, %.1f, %.1f %.1f) instead of (%.1f, %.1f, %.1f %.1f)'
                      % (shape.Left, shape.Top, shape.Width, shape.Height,
                         bbox[0], bbox[1], bbox[2], bbox[3]))
    if not delete_placeholders:
        empty_filled_placeholders(items)


###############################################################################
def replace_figure(picid=-1, keep_aspect=True, **kwargs):
    """ Delete corresponding image from the slide and add a new one on the same
        place
    """
    Slide = get_active_slide()
    # Get all images
    shapes = []
    for ii in range(Slide.Shapes.Count):
        item = Slide.Shapes.Item(1+ii)
        if item.Type == msoPicture:
            shapes.append(shape)
    # Select required shape
    if len(shapes) < picid:
        raise ValueError('Picture index is out of range')
    else:
        shape = shapes[picid]
    # Save position
    pos = [shape.Left, shape.Top, shape.Width, shape.Height]
    # Delete
    shape.Delete()
    # And add a new one
    add_figure(bbox=pos, keep_aspect=keep_aspect, **kwargs)
