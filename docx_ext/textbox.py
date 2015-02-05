# encoding: utf-8
from __future__ import absolute_import, unicode_literals
from docx.oxml import register_element_cls, CT_RPr
from docx.oxml.xmlchemy import ZeroOrOne, BaseOxmlElement, ZeroOrMore
from docx.shared import Parented
from docx.table import Table
from docx.text import Paragraph
from docx.oxml.ns import nsmap


__author__ = 'bluec0re'

nsmap['mc'] = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
nsmap['wps'] = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
nsmap['v'] = 'urn:schemas-microsoft-com:vml'


class CT_TxbxContent(BaseOxmlElement):
    """
    ``<w:rStyle>`` element, containing the properties for a style.
    """
    p = ZeroOrMore('w:p')
    tbl = ZeroOrMore('w:tbl')


register_element_cls('w:txbxContent', CT_TxbxContent)


class Textbox(Parented):
    def __init__(self, parent, element):
        super(Textbox, self).__init__(parent)
        self._tb = element

    @property
    def paragraphs(self):
        return [Paragraph(parent=self, p=p) for p in self._tb.p_lst]

    @property
    def tables(self):
        return [Table(parent=self, tbl=t) for t in self._tb.tbl_lst]
