from __future__ import absolute_import, print_function, unicode_literals

import jinja2
from lxml import etree, html
from lxml.html import clean
import logging
import sys
import zipfile
from preprocess import preprocess
import json
from cgi import escape


__author__ = 'bluec0re'


log = logging.getLogger(__name__)


def transform_html(root, init=False, default_style=None):
    bold = root.tag == 'strong' or 'bold' in root.attrib.get('style', '')
    italic = root.tag == 'i' or 'italic' in root.attrib.get('style', '')
    if root.tag == 'li':
        default_style = root.getparent().get('class', 'ListParagraph')
    if root.tag == 'pre':
        default_style = 'poc'
    style = root.attrib.get('class', default_style)
    new_paragraph = root.tag in ('p', 'h1', 'h2', 'h3', 'h4', 'li', 'pre', 'br') and not init or style
    new_r = not init or new_paragraph or bold or italic

    result = ''
    if new_r:
        result += '</w:t></w:r>'

    if new_paragraph:
        result += '</w:p>'
        result += '<w:p><w:pPr>'
        if style:
            result += '<w:pStyle w:val="%s"/>' % escape(style)
        if root.tag == 'li':
            result += '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr>'
        result += '<w:rPr></w:rPr>'
        result += '</w:pPr>'

    if new_r:
        result += '<w:r><w:rPr>'
        if bold:
            result += '<w:b />'
        result += '</w:rPr><w:t>'

    if root.text is not None:
        result += escape(root.text).strip()

    for child in root.getchildren():
        result += transform_html(child, default_style=style)

    if root.tail is not None:
        result += escape(root.tail).strip()
    return result


def preprocess_html(context):
    if isinstance(context, dict):
        for key, value in context.items():
            context[key] = preprocess_html(value)
        return context
    elif isinstance(context, list):
        return [preprocess_html(v) for v in context]
    elif isinstance(context, tuple):
        return (preprocess_html(v) for v in context)
    elif isinstance(context, (str, unicode)):
        # clean html first
        cleaner = clean.Cleaner()
        cleaner.safe_attrs_only = True
        cleaner.safe_attrs = ('style', 'class')
        cleaner.allow_tags = ('p', 'a', 'br', 'span', 'strong', 'h1', 'h2', 'h3', 'h4', 'i', 'ul', 'li', 'br', 'pre')
        cleaner.remove_unknown_tags = False
        h = cleaner.clean_html(context)
        h = html.fromstring(h)

        # transform to docx code
        if h.find('p') is not None or h.find('span') is not None or\
                        h.find('strong') is not None or h.find('a') is not None:
            value = transform_html(h, True)
        else:
            # remove enclosing tag
            roottag = h.tag
            value = etree.tostring(h)
            value = value[len(roottag) + 2:-(len(roottag)+3)]
        return value
    else:
        return context


def render(doc, context, debug=False):
    if isinstance(doc, etree._Element):
        doc = etree.tostring(doc,
                             encoding='utf-8',
                             xml_declaration=True,
                             standalone=True).decode('utf-8')

    template = jinja2.Template(doc)
    context = preprocess_html(context)
    if debug:
        doc = template.render(**context).encode('utf-8')
        with open('templated.xml', 'w') as fp:
            fp.write(doc)
        doc = etree.XML(doc)
    else:
        doc = etree.XML(template.render(**context).encode('utf-8'))

    # cleanup control nodes
    for el in doc.xpath('//*[@is_control="true"]'):
        par = el.getparent()
        par.remove(el)
        if par.find('w:r/w:t', par.nsmap) is None:
            par.getparent().remove(par)

    return etree.tostring(doc,
                          encoding='utf-8',
                          xml_declaration=True,
                          standalone=True)


def main(preproc=True):
    zipin = zipfile.ZipFile(sys.argv[1])

    if preproc:
        doc = preprocess(zipin.open('word/document.xml'), debug=True)
    else:
        doc = zipin.read('word/document.xml').decode('utf-8')

    processed_doc = render(doc, json.load(sys.stdin))

    print(processed_doc)

    target = 'Processed_' + sys.argv[1]
    outzip = zipfile.ZipFile(target, "w")
    for fileinfo in zipin.infolist():
        if fileinfo.filename != 'word/document.xml':
            outzip.writestr(fileinfo, zipin.read(fileinfo))
        else:
            outzip.writestr('word/document.xml', processed_doc)


if __name__ == '__main__':
    logging.basicConfig(level='DEBUG')
    main()
