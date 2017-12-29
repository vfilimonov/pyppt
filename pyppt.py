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
import os

__version__ = '0.1'
__author__ = 'Vladimir Filimonov'
__email__ = 'vladimir.a.filimonov@gmail.com'

###############################################################################
# Some constants from MSDN MS Office reference
# See also VBA reference: https://github.com/OfficeDev/VBA-content
###############################################################################
# MsoShapeType: https://msdn.microsoft.com/en-us/library/aa432678.aspx
msoShapeType = {'msoShapeTypeMixed': -2,
                'msoAutoShape': 1,
                'msoCallout': 2,
                'msoChart': 3,
                'msoComment': 4,
                'msoFreeform': 5,
                'msoGroup': 6,
                'msoEmbeddedOLEObject': 7,
                'msoFormControl': 8,
                'msoLine': 9,
                'msoLinkedOLEObject': 10,
                'msoLinkedPicture': 11,
                'msoOLEControlObject': 12,
                'msoPicture': 13,
                'msoPlaceholder': 14,
                'msoTextEffect': 15,
                'msoMedia': 16,
                'msoTextBox': 17,
                'msoScriptAnchor': 18,
                'msoTable': 19,
                'msoCanvas': 20,
                'msoDiagram': 21,
                'msoInk': 22,
                'msoInkComment': 23,
                'msoIgxGraphic': 24}

# PpPlaceholderType: https://msdn.microsoft.com/en-us/VBA/PowerPoint-VBA/articles/ppplaceholdertype-enumeration-powerpoint
ppPlaceholderType = {'ppPlaceholderMixed': -2,
                     'ppPlaceholderTitle': 1,
                     'ppPlaceholderBody': 2,
                     'ppPlaceholderCenterTitle': 3,
                     'ppPlaceholderSubtitle': 4,
                     'ppPlaceholderVerticalTitle': 5,
                     'ppPlaceholderVerticalBody': 6,
                     'ppPlaceholderObject': 7,
                     'ppPlaceholderChart': 8,
                     'ppPlaceholderBitmap': 9,
                     'ppPlaceholderMediaClip': 10,
                     'ppPlaceholderOrgChart': 11,
                     'ppPlaceholderTable': 12,
                     'ppPlaceholderSlideNumber': 13,
                     'ppPlaceholderHeader': 14,
                     'ppPlaceholderFooter': 15,
                     'ppPlaceholderDate': 16,
                     'ppPlaceholderVerticalObject': 17,
                     'ppPlaceholderPicture': 18}


pp_titles = [ppPlaceholderType['ppPlaceholder' + _] for _ in ('Title', 'Subtitle', 'Body')]
pp_pictures = [ppPlaceholderType['ppPlaceholder' + _] for _ in ('Object', 'Bitmap', 'Picture')]

# MsoZOrderCmd: https://msdn.microsoft.com/en-us/library/aa432726.aspx
msoZOrderCmd = {'msoBringToFront': 0,
                'msoSendToBack': 1,
                'msoBringForward': 2,
                'msoSendBackward': 3,
                'msoBringInFrontOfText': 4,
                'msoSendBehindText': 5}

# For reverse lookup
msoShapeTypeInt = {v: k for k, v in msoShapeType.items()}
ppPlaceholderTypeInt = {v: k for k, v in ppPlaceholderType.items()}
msoZOrderCmdInt = {v: k for k, v in msoZOrderCmd.items()}

# Temporary text to be filled in empty placeholders
_TEMPTEXT = '--TO-BE-REMOVED--'

###############################################################################
# Presets for the image positions in the format [x, y, width, height]
# If all values are in the range [0, 1], then they are treated as fractions
# from the slide width and height, otherwise, they will be treated as values
# in pixels
###############################################################################
presets = {'Center': [0.0415, 0.227, 0.917, 0.716],
           'CenterL': [0.0415, 0.153, 0.917, 0.716],
           'CenterXL': [0.0415, 0.049, 0.917, 0.888],
           'CenterXXL': [0, 0, 1., 1.],
           'Full': [0, 0, 1., 1.]}


###############################################################################
# General methods for accessing presentation and slides
###############################################################################
def _temp_fname():
    """ Return a name of a temporary file """
    f = tempfile.NamedTemporaryFile(delete=False)
    f.close()
    name = f.name
    os.remove(name)
    return name + '.png'


def _get_application():
    """ Get reference to PowerPoint application """
    Application = client.Dispatch('PowerPoint.Application')
    # Make it visible
    Application.Visible = True
    return Application


def _get_active_presentation():
    """ Get reference to active presentation """
    return _get_application().ActivePresentation


def _get_slide(slide_no=None):
    """ Get reference to a slide with a given number (indexing starts from 1).
        If number is not specified, then the active slide is returned.
    """
    if slide_no is None:
        return _get_application().ActiveWindow.View.Slide
    else:
        Presentation = _get_active_presentation()
        return Presentation.Slides[slide_no - 1]


###############################################################################
# Convenience methods for working with placeholders
###############################################################################
def _shapes(Slide, types=None):
    """ Return all Shapes from the given slide.
        If types are provided, then Shapes of only given types will be returned.
    """
    if Slide is None or isinstance(Slide, (int, long)):
        # Infer from the number
        Slide = _get_slide(Slide)
    shapes = [Slide.Shapes.Item(1+ii) for ii in range(Slide.Shapes.Count)]
    if types is not None:
        types = [msoShapeType[_] for _ in types]
        shapes = [s for s in shapes if s.Type in types]
    return shapes


def _placeholders(Slide):
    """ Wrapper for the _shapes to return only placeholders. """
    return _shapes(Slide, ['msoPlaceholder'])


def _placeholders_pictures(Slide, empty=False):
    """ List of all placeholders for pictures.
        If empty is True - keep only empty placeholders
    """
    pics = [p for p in _placeholders(Slide)
            if p.PlaceholderFormat.type in pp_pictures]
    if empty:
        pics = [p for p in pics if _is_placeholder_empty(p)]
    return pics


def _has_textframe(obj):
    """ Check if placeholder has TextFrame """
    return hasattr(obj, 'TextFrame') and hasattr(obj.TextFrame, 'TextRange')


def _is_placeholder_empty(obj):
    """ Check if placeholder is empty """
    if obj.PlaceholderFormat.ContainedType == msoShapeType['msoAutoShape']:
        # Either a content/text or another type with no content
        if _has_textframe(obj):
            if obj.TextFrame.TextRange.Length == 0:
                return True
        else:
            return True
    return False


def _empty_placeholders(Slide):
    """ Return a list of empty placeholders """
    return [s for s in _placeholders(Slide) if _is_placeholder_empty(s)]


###############################################################################
# ...treating empty placeholders
###############################################################################
def _fill_empty_placeholders(Slide):
    """ Dirty hack: fill all empty placeholders with some text.
        Returns a list of objects that were filled, so then the text could be
        cleared from them (see empty_placeholders()).
        If we don't do this - when we insert a figure, placeholder will disappear.
    """
    filled = []
    for p in _empty_placeholders(Slide):
        if p.PlaceholderFormat.type not in pp_titles:
            if _has_textframe(p):
                p.TextFrame.TextRange.Text = _TEMPTEXT
                filled.append(p)
    return filled


def _revert_filled_placeholders(items):
    """ Remove text from all placeholders that were filled by
        _fill_empty_placeholders()
    """
    for item in items:
        item.TextFrame.TextRange.Text = ''


def _delete_empty_placeholders(Slide):
    """ Delete all empty placeholders except Title and Subtitle """
    # we're going ro remove => iterate in reverse order
    for p in _empty_placeholders(Slide)[::-1]:
        if p.PlaceholderFormat.type not in pp_titles:
            p.delete()


###############################################################################
# Titles and Subtitles
###############################################################################
def title_to_front(slide_no=None):
    """ Bring title and subtitle to front """
    titles = [p for p in _placeholders(_get_slide(slide_no))
              if p.PlaceholderFormat.type in pp_titles]
    for item in titles:
        item.ZOrder(msoZOrderCmd['msoBringToFront'])


def set_title(title, slide_no=None):
    """ Set title for the slide (active or of a given number).
        If slide contain multiple Placeholder/Title objects, only first one is set.
    """
    for p in _placeholders(_get_slide(slide_no)):
        if p.PlaceholderFormat.type == ppPlaceholderType['ppPlaceholderTitle']:
            p.TextFrame.TextRange.Text = title
            return
    warnings.warn('No title placeholders were found on the given slide')


def set_subtitle(subtitle, slide_no=None):
    """ Set title for the slide (active or of a given number).
        If slide contain multiple Placeholder/Title objects, only first one is set.
    """
    for p in _placeholders(_get_slide(slide_no)):
        if p.PlaceholderFormat.type == ppPlaceholderType['ppPlaceholderSubtitle']:
            p.TextFrame.TextRange.Text = subtitle
            return
    warnings.warn('No subtitle placeholders were found on the given slide')


def add_slide(slide_no=None, layout_as=None):
    """ Add slide after slide number "slide_no" with the layout as in the slide
        number "layout_as".
        If "slide_no" is None, new slide will be added after the active one.
        If "layout_as" is None, new slide will have layout as the active one.
        Returns the number of the added slide.
    """
    if slide_no is None:
        slide_no = _get_slide().SlideNumber
    if layout_as is None:
        layout_as = slide_no
    Presentation = _get_active_presentation()
    pptLayout = Presentation.Slides[layout_as - 1].CustomLayout
    Slide = Presentation.Slides.AddSlide(slide_no + 1, pptLayout)
    return Slide.SlideNumber


###############################################################################
# Extracting metadata
###############################################################################
def get_shape_positions(slide_no=None):
    """ Get positions of all shapes in the slide.
        Return list of lists of the format [x, y, w, h, type].
    """
    return [[item.Left, item.Top, item.Width, item.Height, item.Type]
            for item in _shapes(_get_slide(slide_no))]


def get_image_positions(slide_no=None, asarray=True, decimals=1):
    """ Get positions of all images in the slide. If necessary, rounds
        coordinates to a given decimals (if "decimals" is not None)
        Return list of lists of the format [x, y, w, h].
    """
    positions = get_shape_positions(slide_no)
    # Keep only images
    positions = [p[:-1] for p in positions if p[-1] == msoShapeType['msoPicture']]
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
# Core functionality
###############################################################################
def _parse_bbox(bbox, keep_aspect=True):
    """ Human-readable bbox-dimensions."""
    if isinstance(bbox, str):  # Use preset
        bbox = presets[bbox]
        if all([0 <= _ <= 1 for _ in bbox]):
            W, H = get_slide_dimensions()
            bbox = [bbox[0] * W, bbox[1] * H, bbox[2] * W, bbox[3] * H]

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
def add_figure(bbox=None, slide_no=None, keep_aspect=True, tight=True,
               delete_placeholders=True, **kwargs):
    """ Add current figure to the active slide (or a slide with a given number).

        Parameters:
            bbox - Bounding box for the image in the format:
                    - None - the first empty image placeholder will be used, if
                             no such placeholders are found, then the 'Center'
                             value will be used.
                    - list of coordinates [x, y, width, height]
                    - string: 'Center', 'Left', 'Right', 'TopLeft', 'TopRight',
                      'BottomLeft', 'BottomRight', 'CenterL', 'CenterXL', 'Full'
                      based on the presets, that could be modified
            slide_no - number of the slide (stating from 1), where to add image.
                       if not specified (None), active slide will be used.
            keep_aspect - if True, then the aspect ratio of the image will be
                          preserved, otherwise the image will shrink to fit bbox.
            tight - if True, then tight_layout() will be used
            delete_placeholders - if True, then all placeholders will be deleted
            **kwargs - to be passed to plt.savefig()
    """
    # Save the figure to png in temporary directory
    fname = _temp_fname()
    if tight:
        # Usually is an overkill, but is needed sometimes...
        plt.tight_layout()
        plt.savefig(fname, bbox_inches='tight', **kwargs)
    else:
        plt.savefig(fname, **kwargs)

    Slide = _get_slide(slide_no)

    # Parse bbox name if necessary
    use_placeholder = False
    if bbox is None:
        # Try to get position of the first empty placeholder for pictures
        pictures = _placeholders_pictures(Slide, empty=True)
        try:
            item = pictures[0]
            bbox = [item.Left, item.Top, item.Width, item.Height]
            use_placeholder = True
        except IndexError:
            # If no placholders: use 'Center'
            bbox = 'Center'
    bbox = _parse_bbox(bbox, keep_aspect=keep_aspect)

    # Now insert to PowerPoint
    if not use_placeholder:
        if delete_placeholders:
            _delete_empty_placeholders(Slide)
        else:
            items = _fill_empty_placeholders(Slide)
    shape = Slide.Shapes.AddPicture(FileName=fname, LinkToFile=False,
                                    SaveWithDocument=True, Left=bbox[0],
                                    Top=bbox[1], Width=bbox[2], Height=bbox[3])
    filled_bbox = [shape.Left, shape.Top, shape.Width, shape.Height]

    # Check if the bbox is correctly filled.
    if np.max(np.abs(np.array(bbox)-np.array(filled_bbox))) > 0.1:
        # Should never happen...
        warnings.warn('BBox of the inserted figure was not respected: '
                      '(%.1f, %.1f, %.1f %.1f) instead of (%.1f, %.1f, %.1f %.1f)'
                      % (shape.Left, shape.Top, shape.Width, shape.Height,
                         bbox[0], bbox[1], bbox[2], bbox[3]))

    # Clean-up
    if not use_placeholder and not delete_placeholders:
        _revert_filled_placeholders(items)
    os.remove(fname)


###############################################################################
def replace_figure(pic_no=-1, slide_no=None, keep_aspect=True, **kwargs):
    """ Delete an image from the slide and add a new one on the same place """
    Slide = _get_slide(slide_no)
    # Get all images
    shapes = []
    for ii in range(Slide.Shapes.Count):
        item = Slide.Shapes.Item(1+ii)
        if item.Type == msoShapeType['msoPicture']:
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
