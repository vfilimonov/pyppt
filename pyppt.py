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
ppPlaceholderObject = 7
ppPlaceholderBitmap = 9
ppPlaceholderPicture = 18
pp_titles = (ppPlaceholderTitle, ppPlaceholderSubtitle, ppPlaceholderBody)
pp_pictures = (ppPlaceholderObject, ppPlaceholderBitmap, ppPlaceholderPicture)
# MsoZOrderCmd: https://msdn.microsoft.com/en-us/library/aa432726.aspx
msoBringToFront = 0
# Temporary text to be filled in empty placeholders
TEMPTEXT = '--TO-BE-REMOVED--'


###############################################################################
def _temp_fname():
    """ Return a name of a temporary file """
    f = tempfile.NamedTemporaryFile(delete=False)
    f.close()
    return f.name + '.png'


def _get_application():
    """ Get reference to PowerPoint application """
    Application = client.Dispatch('PowerPoint.Application')
    # Make it visible
    Application.Visible = True
    return Application


def _get_active_presentation():
    """ Get reference to active presentation """
    return _get_application().ActivePresentation


def _get_slide(index=None):
    """ Get reference to a slide with a given number (indexing starts from 1).
        If number is not specified, then the active slide is returned.
    """
    if index is None:
        return _get_application().ActiveWindow.View.Slide
    else:
        Presentation = _get_active_presentation()
        return Presentation.Slides[index]


def _shapes(Slide, types=None, reverse=False):
    """ Generator: yields Shapes from the given slide.
        If types are provided, then Shapes of only given types will be yielded.
        If reverse, then shapes will be yielded in reverse order.
    """
    rng = range(Slide.Shapes.Count)
    if reverse:
        rng = rng[::-1]
    for ii in rng:
        item = Slide.Shapes.Item(1+ii)
        if types is None:
            yield item
        elif item.Type in types:
            yield item


def _placeholders(Slide, reverse=False):
    """ Wrapper for the _shapes to yield only placeholders. """
    for item in _shapes(Slide, [msoPlaceholder], reverse=reverse):
        yield item


###############################################################################
def _fill_empty_placeholders(Slide, text='Temp'):
    """ Dirty hack: fill all empty placeholders with some text.
        Returns a list of objects that were filled, so then the text could be
        cleared from them (see empty_placeholders()).
        If we don't do this - when we insert a figure, it will take place of
        the placeholder, rather than using specified coordinates.
    """
    placeholders = [p for p in _placeholders(Slide)
                    if p.PlaceholderFormat.type not in pp_titles]
    filled = []
    for item in placeholders:
        try:
            if item.TextFrame.TextRange.Length == 0:  # if empty
                item.TextFrame.TextRange.Text = TEMPTEXT
                filled.append(item)
        except:
            # Something happened...
            # Most likely this is not the one we want to fill
            pass
    return filled


def _empty_filled_placeholders(items):
    """ Remove text from all placeholders that were filled by
        _fill_empty_placeholders()
    """
    for item in items:
        item.TextFrame.TextRange.Text = ''


def _delete_empty_placeholders(Slide):
    """ Delete all empty placeholders except Title and Subtitle """
    # we're going ro remove => iterate in reverse order
    placeholders = [p for p in _placeholders(Slide)
                    if p.PlaceholderFormat.type not in pp_titles]
    for item in placeholders[::-1]:
        try:
            if item.TextFrame.TextRange.Length == 0:  # if empty
                item.delete()
        except:
            # Something happened...
            # Most likely this is not the one we want to delete
            pass


def title_to_front(slide_no=None):
    """ Bring title and subtitle to front """
    Slide = _get_slide(slide_no)
    titles = [p for p in _placeholders(Slide)
              if p.PlaceholderFormat.type in pp_titles]
    for item in titles:
        item.ZOrder(msoBringToFront)


###############################################################################
def get_shape_positions(slide_no=None):
    """ Get positions of all shapes in the slide.
        Return list of lists of the format [x, y, w, h, type].
    """
    for item in _shapes(_get_slide(slide_no), reverse=False):
        positions.append([item.Left, item.Top, item.Width, item.Height, item.Type])
    return positions


def get_image_positions(slide_no=None, asarray=True, decimals=1):
    """ Get positions of all images in the slide. If necessary, rounds
        coordinates to a given decimals (if "decimals" is not None)
        Return list of lists of the format [x, y, w, h].
    """
    positions = get_shape_positions(slide_no)
    # Keep only images
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
        Presentation = _get_active_presentation()
    return (Presentation.PageSetup.SlideWidth,
            Presentation.PageSetup.SlideHeight)


def get_notes(Presentation=None):
    """ Extract notes for all slides from the presentation """
    if Presentation is None:
        Presentation = _get_active_presentation()
    Slides = Presentation.Slides
    notes = []
    for ii in range(len(Slides)):
        notes.append(Slides[ii].NotesPage.Shapes.Placeholders[2]
                               .TextFrame.TextRange.Text)
    return notes


###############################################################################
def _parse_bbox(bbox, Slide, keep_aspect=True):
    """ Human-readable bbox-dimensions"""
    if bbox is None:
        pass

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
    return bbox


###############################################################################
def add_figure(bbox=None, slide_no=None, keep_aspect=True,
               delete_placeholders=True, bbox_inches='tight', **kwargs):
    """ Add current figure to the active slide (or slide with a given number).

    """
    # Save the figure to png
    fname = _temp_fname()
    if bbox_inches == 'tight':
        # Usually is an overkill, but is needed sometimes...
        plt.tight_layout()
    plt.savefig(fname, bbox_inches=bbox_inches, **kwargs)

    Slide = _get_slide(slide_no)

    # Parse bbox name if necessary
    bbox = _parse_bbox(bbox, Slide, keep_aspect=keep_aspect)

    # Now insert to PowerPoint
    if delete_placeholders:
        _delete_empty_placeholders(Slide)
    else:
        items = _fill_empty_placeholders(Slide)
    shape = Slide.Shapes.AddPicture(FileName=fname, LinkToFile=False,
                                    SaveWithDocument=True, Left=bbox[0],
                                    Top=bbox[1], Width=bbox[2], Height=bbox[3])
    filled_bbox = [shape.Left, shape.Top, shape.Width, shape.Height]
    # Check if the bbox is correctly filled.
    # Should happen always...
    if np.max(np.abs(np.array(bbox)-np.array(filled_bbox))) > 0.1:
        warnings.warn('BBox of the inserted figure was not respected: '
                      '(%.1f, %.1f, %.1f %.1f) instead of (%.1f, %.1f, %.1f %.1f)'
                      % (shape.Left, shape.Top, shape.Width, shape.Height,
                         bbox[0], bbox[1], bbox[2], bbox[3]))
    if not delete_placeholders:
        _empty_filled_placeholders(items)


###############################################################################
def replace_figure(pic_no=-1, slide_no=None, keep_aspect=True, **kwargs):
    """ Delete corresponding image from the slide and add a new one on the same
        place
    """
    Slide = _get_slide(slide_no)
    # Get all images
    shapes = []
    for ii in range(Slide.Shapes.Count):
        item = Slide.Shapes.Item(1+ii)
        if item.Type == msoPicture:
            shapes.append(shape)
    # Select required shape
    if len(shapes) < pic_no:
        raise ValueError('Picture index is out of range')
    else:
        shape = shapes[pic_no]
    # Save position
    pos = [shape.Left, shape.Top, shape.Width, shape.Height]
    # Delete
    shape.Delete()
    # And add a new one
    add_figure(bbox=pos, keep_aspect=keep_aspect, **kwargs)
