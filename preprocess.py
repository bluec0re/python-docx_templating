from __future__ import print_function, unicode_literals
import shlex

__author__ = 'bluec0re'

import zipfile
import sys
from lxml import etree
import logging
import re

log = logging.getLogger(__name__)


is_control_field = re.compile(r'^\s*\{%\s*(end)?(for|if)')


def isTag(element, tagname):
    if element.prefix:
        return element.tag == '{%s}%s' % (element.nsmap[element.prefix], tagname)
    else:
        return element.tag == tagname


def get_attrib(element, attribname):
    if ':' in attribname:
        prefix, attribname = attribname.split(':', 1)
        return element.attrib.get('{%s}%s' % (element.nsmap[prefix], attribname))
    return element.attrib.get(attribname)


def main():
    zipin = zipfile.ZipFile(sys.argv[1])

    doc = preprocess(zipin.open('word/document.xml'), debug=True)

    target = 'Preprocessed_' + sys.argv[1]
    outzip = zipfile.ZipFile(target, "w")
    for fileinfo in zipin.infolist():
        if fileinfo.filename != 'word/document.xml':
            outzip.writestr(fileinfo, zipin.read(fileinfo))
        else:
            outzip.writestr('word/document.xml', etree.tostring(doc,
                                                                encoding='utf-8',
                                                                xml_declaration=True,
                                                                standalone=True))


def parse_field(field):
    lex = shlex.shlex(field, posix=True)
    lex.whitespace_split = True
    lex.commenters = ''
    lex.escape = ''
    tokens = list(lex)
    #     0         1      2     3
    # MERGEFIELD "<entry>" \* FORMAT
    return tokens[1]


def update_text(t, controls, mergefield):
    if '\u00AB' in t.text:
        t.text = t.text.replace('\u00AB', '').replace('\u00BB', '').replace('$', '')
        if '#' in t.text:  # trying autoconvert
            t.text, n = re.subn(r'#foreach\((.+?)\)', 'for \\1', t.text, 0, re.I)
            if n:
                controls.append('endfor')
            else:
                t.text, n = re.subn(r'#if\((.+?)\)', 'if \\1', t.text, 0, re.I)
                if n:
                    controls.append('endif')
                elif 'end' in t.text and controls:
                    t.text = t.text.replace(r'#end', controls.pop(-1), t.text)
                else:
                    log.warning("Unkown directive %s", t.text)
            t.text = t.text.replace('foreach.isFirst', 'loop.first')\
                           .replace('foreach.isLast', 'loop.last')\
                           .replace('foreach.hasNext', 'not loop.last')
            t.text = "{%% %s %%}" % t.text
        elif '{' not in t.text:
            t.text = "{{ %s }}" % t.text

        # take mergefield content
        if t.text != mergefield:
            t.text = mergefield

        log.debug("Field text: %s", t.text)

        # mark control fields
        if is_control_field.search(t.text):
            t.getparent().attrib['is_control'] = 'true'
    else:
        log.warning("No marker: %s", t.text)


def search_fldChar(start, method, type, controls, mergefield):
    stack = [getattr(start, method)()]
    result = None
    text = None
    while stack:
        el = stack.pop(0)
        if el is None:
            raise ValueError("No %s field found" % type)
        fldChar = el.find('w:fldChar', el.nsmap)
        if fldChar is not None and get_attrib(fldChar, 'w:fldCharType') == type:
            result = el
            break

        t = el.find('w:t', el.nsmap)
        if t is not None:
            text = el
            update_text(t, controls, mergefield)
        stack.append(getattr(el, method)())
    return result, text


def remove_unneeded(next, end, text):
    finish = False
    tmpparent = next.getparent()
    while not finish:
        if next == end:
            finish = True
        tmp = next.getnext()
        if next != text:
            tmpparent.remove(next)
        next = tmp


def parse_complex_fields(doc):
    for field in doc.xpath('//w:instrText[contains(text(), "MERGEFIELD")]', namespaces=doc.nsmap):
        log.debug('Complex Field %s found', field.text)
        parent = field.getparent()
        assert isTag(parent, 'r'), "%s is not r element" % parent

        mergefield = field.text

        try:
            mergefield = parse_field(mergefield)
        except:
            log.warning("Splitted field")
            mergefield += parent.getnext().find('w:instrText', field.nsmap).text
            log.debug("Field %s", mergefield)
            mergefield = parse_field(mergefield)

        # get <w:fldChar w:fldCharType="begin"/>
        controls = []
        start, text = search_fldChar(parent, 'getprevious', 'begin', controls, mergefield)

        # get <w:fldChar w:fldCharType="end"/>
        end, text = search_fldChar(parent, 'getnext', 'end', controls, mergefield)

        # remove unneded
        remove_unneeded(start, end, text)


def parse_simple_fields(doc):
    for field in doc.xpath('//w:fldSimple[contains(@w:instr, "MERGEFIELD")]', namespaces=doc.nsmap):
        instr = get_attrib(field, 'w:instr')
        parent = field.getparent()
        log.debug('Simple Field %s found', instr)

        # parse instruction
        mergefield = parse_field(instr)

        # copy content
        prev = field.getprevious()
        if prev is not None:
            for child in field.getchildren():
                t = child.findall('.//w:t', child.nsmap)[0]
                update_text(t, [], mergefield)
                prev.addnext(child)
        else:
            log.warning("Field %s has no prev", instr)
            for child in field.getchildren():
                t = child.findall('.//w:t', child.nsmap)[0]
                update_text(t, [], mergefield)
                parent.append(child)

        # remove old field
        parent.remove(field)


def preprocess(document, debug=False):
    if hasattr(document, 'read'):
        document = document.read()

    if debug:
        with open('orig.xml', 'w') as fp:
            fp.write(document)

    doc = etree.XML(document)

    if debug:
        with open('parsed.xml', 'w') as fp:
            fp.write(etree.tostring(doc,
                                    encoding='utf-8',
                                    pretty_print=True,
                                    xml_declaration=True,
                                    standalone=True))

    # fldChar
    parse_complex_fields(doc)

    # fldSimple
    parse_simple_fields(doc)

    if debug:
        with open('new.xml', 'w') as fp:
            fp.write(etree.tostring(doc,
                                    encoding='utf-8',
                                    pretty_print=True,
                                    xml_declaration=True,
                                    standalone=True))

    return doc

if __name__ == '__main__':
    logging.basicConfig(level='DEBUG')
    main()
