# encoding: utf-8
from __future__ import absolute_import, unicode_literals

import re
import logging
import os

__author__ = 'bluec0re'

log = logging.getLogger(__name__)


class Default(object):
    def __repr__(self):
        attrs = ', '.join(
                    '%s=%s' % (key, repr(getattr(self, key))) for key in self.__dict__
                    if not key.startswith('_'))
        return '{}({})'.format(type(self).__name__, attrs)

    def __str__(self):
        return repr(self)

    def __unicode__(self):
        s = str(self)
        return s
        return str(self).decode('utf-8')

    def __eq__(self, other):
        try:
            return all(getattr(self, key) == getattr(other, key)
                       for key in self.__dict__ if not key.startswith('_'))
        except AttributeError:
            return False


def make_relative(path, start):
    log.debug('%s - %s', path, start)
    if path == start or start is None:
        result = './'
    else:
        newpath = os.path.relpath(path, start)
        if newpath.startswith('../'):
            m = re.search(r'\.\./([^/]+)\[(\d+)\]', newpath)

            tag = m.group(1)
            newoffset = int(m.group(2))
            m = re.search(r'/%s\[(\d+)\]' % re.escape(tag), start)

            oldoffset = int(m.group(1))
            result = newpath.replace('../%s[%d]' % (tag, newoffset), '../%s[%d]' % (tag, newoffset - oldoffset))
        else:
            result = './' + newpath
    log.debug(' => %s', result)
    return result


def make_abs(path, root):
    log.debug('%s + %s', path, root)
    if path == './':
        result = root
    elif path.startswith('./'):
        result = root + '/' + path[2:]
    elif path.startswith('../'):
        m = re.search('\.\./([^/]+)\[(-?\d+)\]', path)
        tag = m.group(1)
        offset = int(m.group(2))
        path = path.replace(m.group(0), '')
        while path.startswith('../'):
            root = root.rsplit('/', 1)[0]
            path = path[3:]
        if offset > 0:
            if tag in root:
                result = root + '/following-sibling::%s[%d]' % (tag, offset) + path
            else:
                result = root + '/%s[%d]' % (tag, offset) + path
        else:
            m = re.search(r'%s\[(\d+)\]' % re.escape(tag), root)
            oldoffset = int(m.group(1))
            newoffset = oldoffset + offset
            result = root.replace(m.group(0), '') + ('%s[%d]' % (tag, newoffset)) + path
    else:
        result = root + '/' + path

    log.debug(' => %s', result)
    return result
