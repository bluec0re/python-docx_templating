# encoding: utf-8
from __future__ import absolute_import, unicode_literals
from docx import Document
from docx.parts.document import DocumentPart
from docx.oxml.ns import qn
from docx_ext.textbox import Textbox


__author__ = 'bluec0re'


# noinspection PyProtectedMember
def _styles(self):
    styles = {}
    for style in self.styles_part.styles._styles_elm.style_lst:
        style_id = style.attrib[qn('w:styleId')]
        style_name = style.find(qn('w:name')).attrib[qn('w:val')]
        styles[style_id] = style_name
    return styles

Document.styles = property(_styles)


def _textboxes(self):
    tboxes = []
    for tb in self._element.body.xpath('//w:txbxContent'):
        tboxes.append(Textbox(parent=self, element=tb))
    return tboxes

DocumentPart.textboxes = property(_textboxes)


def _textboxes(self):
    return self._document_part.textboxes

Document.textboxes = property(_textboxes)
