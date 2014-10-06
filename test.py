# encoding: utf-8
from __future__ import absolute_import, unicode_literals

from pprint import pprint
import logging
from PIL import Image

from docx import Document
# experiments with .docm
#from docx.parts.document import DocumentPart
#from docx.opc.package import PartFactory
#from docx.opc.constants import CONTENT_TYPE

from parser import Context, gen_tree

import docx_ext

docx_ext.init()

log = logging.getLogger(__name__)

__author__ = 'bluec0re'


def img_factory(path):
    return Image.open(path)

docx_ext.register_image_factory(img_factory)


def main():
    global_vars = Context({
        'items': [
            {'name': 'Item A', 'description': 'Its just item A'},
            {'name': 'Item B', 'description': 'Its just item B.\nBut still more than A'},
            {'name': 'Item C', 'description': '<p>This time it\'s item <span style="color: #00ff00">C</span>!</p>'
                                              '<p>With this nice image: <img src="lena.png" alt="Another picture!"></p>'
                                              '<p>It contains <b>bold</b> and <i>italic</i> Texts and knows'
                                              '<h3>headings</h3>'
                                              '</p>'
                                              '<p class="Title">And Titles</p>'}
        ],
        'author': 'Me, who else?',
        'logo': Image.open('lena.png')
    })
    global_vars.variables['logo'].caption = 'Great picture'

    # experiments with .docm
    #CONTENT_TYPE.WML_DOCUMENT_MAIN = 'application/vnd.ms-word.document.macroEnabled.main+xml'
    #PartFactory.part_type_for[CONTENT_TYPE.WML_DOCUMENT_MAIN] = DocumentPart
    #doc = Document("HelloField.docm")
    doc = Document("HelloField.docx")

    tree = gen_tree(doc)
    try:
        tree.evaluate(global_vars, allowed_styles=doc.styles.keys())
    finally:
        doc.save("test.docx")

if __name__ == '__main__':
    logging.basicConfig(level='INFO', format='%(levelname)s-%(module)s.%(funcName)s:%(lineno)d: %(message)s')
    main()
