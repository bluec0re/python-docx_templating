# encoding: utf-8
from __future__ import absolute_import, unicode_literals

from docx.oxml import CT_R, CT_String, register_element_cls, CT_RPr, OxmlElement
from docx.oxml.xmlchemy import ZeroOrOne

from docx.text import Run

__author__ = 'bluec0re'


class CT_RPr2(CT_RPr):
    """
    ``<w:rStyle>`` element, containing the properties for a style.
    """
    _color = ZeroOrOne('w:color')

    @property
    def color(self):
        color = self._color
        if color is None:
            return color
        return color.val

    @color.setter
    def color(self, color):
        if color is None:
            self._remove__color()
        elif self._color is None:
            self._add__color(val=color)
        else:
            self._color.val = color


register_element_cls('w:rPr', CT_RPr2)
register_element_cls('w:color', CT_String)


def _r_color_getter(self):
    """
    String contained in w:color element of <w:rStyle> grandchild, or
    |None| if that element is not present.
    """
    rPr = self.rPr
    if rPr is None:
        return None
    return rPr.color


def _r_color_setter(self, color):
    """
    Set the character style of this <w:r> element to *style*. If *style*
    is None, remove the style element.
    """
    rPr = self.get_or_add_rPr()
    rPr.color = color

CT_R.color = property(_r_color_getter, _r_color_setter)


def _add_r_after(self):
    """
    Return a new ``<w:r>`` element inserted directly prior to this one.
    """
    new_r = OxmlElement('w:r')
    self.addnext(new_r)
    return new_r


def _add_r_before(self):
    """
    Return a new ``<w:r>`` element inserted directly prior to this one.
    """
    new_r = OxmlElement('w:r')
    self.addprevious(new_r)
    return new_r

CT_R.add_r_after = _add_r_after
CT_R.add_r_before = _add_r_before


def _run_color_getter(self):
    return self._r.color


def _run_color_setter(self, color):
    self._r.color = color

Run.color = property(_run_color_getter, _run_color_setter)
