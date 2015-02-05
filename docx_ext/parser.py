# encoding: utf-8
from __future__ import absolute_import, unicode_literals

import logging
from copy import deepcopy
import re
import sys

if sys.version > '3':
    unicode = str

from .utils import Default, make_relative, make_abs

log = logging.getLogger(__name__)

__author__ = 'bluec0re'


class ParserException(Default, Exception):
    def __init__(self, msg, obj):
        super(ParserException, self).__init__()
        self.msg = msg
        self.obj = obj

    def __str__(self):
        if isinstance(self.msg, unicode):
            return self.msg.encode('utf-8')
        else:
            return self.msg


class Context(Default):
    def __init__(self, variables=None, parent=None, root=None):
        super(Context, self).__init__()
        self.root = root
        self.variables = variables or {}
        self.parent = parent

    def resolve(self, path, variables=None):
        if variables is None:
            variables = self.variables

        if not isinstance(variables, dict):
            return variables

        if not isinstance(path, list):
            path = path.split('.')

        if len(path) == 1:
            result = variables.get(path[0])
        else:
            result = self.resolve(path[1:], variables.get(path[0]))

        if result is None and self.parent is not None:
            return self.parent.resolve(path)
        else:
            return result


class FieldBased(Default):
    def __init__(self, field):
        super(FieldBased, self).__init__()
        self.field = field


class Container(Default):
    def __init__(self, parent=None):
        super(Container, self).__init__()
        self.childs = []
        self.parent = parent
        self.start = None
        self.end = None

    def evaluate(self, context, base=None, allowed_styles=None):
        log.debug("Evaluating container")
        for c in self.childs:
            c.evaluate(context,
                       base=base,
                       allowed_styles=allowed_styles)
        self.remove_fields()

    def content(self):
        elements = []
        if self.start.end is None:
            return elements

        root = self.start.end.getparent()
        for el in root.itersiblings():
            if el == self.end.start.getparent():
                break

            elements.append(el)

        return elements

    def copy(self):
        return [deepcopy(e) for e in self.content()]

    def remove_fields(self):
        log.debug("Removing start: %s", self.start)
        if self.start is not None:
            log.debug("Parent: %s", self.start.start.getparent())
            self.start.remove()

        log.debug("Removing end: %s", self.end)
        if self.end is not None:
            log.debug("Parent: %s", self.end.start.getparent())
            self.end.remove()

    def remove_content(self):
        log.debug("Removing content of %s", self)
        for el in self.content():
            log.debug("Removing content el: %s", el)
            el.getparent().remove(el)


class Variable(FieldBased):
    def __init__(self, field, path):
        super(Variable, self).__init__(field)
        self.path = path
        self._value = None

    def resolve(self, context):
        return context.resolve(self.path)

    def evaluate(self, context, base=None, allowed_styles=None):
        log.debug("Evaluating variable %s", self.path)
        value = self.resolve(context)
        self.field.replace(value, base, allowed_styles=allowed_styles)


class If(FieldBased, Container):
    def __init__(self, field, parent=None, src=None):
        super(If, self).__init__(field=field)
        self.parent = parent
        self.src = src
        self.start = field

    def evaluate(self, context, base=None, allowed_styles=None):
        log.debug("Evaluating if %s", self.src)
        code = self.src.replace('!', ' not ')

        class ReplaceVars:
            def __init__(self):
                self.locals = {}
                self._num = 0

            def __call__(self, m):
                var = "var%d" % self._num
                self.locals[var] = context.resolve(m.group(1))
                self._num += 1
                return var

        variables = ReplaceVars()
        code = re.sub(r'\$(\S+)', variables, code)

        if eval(code, variables.locals, {}):
            log.debug("If success")
            super(If, self).evaluate(context,
                                     base=base,
                                     allowed_styles=allowed_styles)
        else:
            log.debug("If failed")
            self.remove_content()
            self.remove_fields()


class ForEach(FieldBased, Container):
    def __init__(self, field, parent=None, dest=None, src=None):
        super(ForEach, self).__init__(field=field)
        self.parent = parent
        self.dest = dest
        self.src = src
        self.start = field

    def itervalues(self, context):
        src = context.resolve(self.src)
        if src is None:
            return

        for i, value in enumerate(src):
            yield Context(variables={
                self.dest: value,
                'foreach': {
                    'isFirst': i == 0,
                    'hasNext': i < len(src) - 1,
                    'isLast': i == len(src) - 1
                }
            }, parent=context, root=self.field.start)

    #noinspection PyProtectedMember
    def _build_cache(self, root_xpath):
        for old_child in self.childs:
            cache = {'parent': old_child.field._parent}

            if hasattr(old_child, 'end'):
                cache['e_parent'] = old_child.end._parent
            else:
                cache['e_parent'] = None

            if hasattr(old_child, 'start'):
                cache['s_parent'] = old_child.start._parent
            else:
                cache['s_parent'] = None

            cache['field.xpath_start'] = make_relative(old_child.field.xpath_start, root_xpath)
            cache['field.xpath_end'] = make_relative(old_child.field.xpath_end, root_xpath)

            if hasattr(old_child, 'start') and old_child.start != old_child.field:
                cache['start.xpath_start'] = make_relative(old_child.start.xpath_start, root_xpath)
                cache['start.xpath_end'] = make_relative(old_child.start.xpath_end, root_xpath)

            if hasattr(old_child, 'end'):
                cache['end.xpath_start'] = make_relative(old_child.end.xpath_start, root_xpath)
                cache['end.xpath_end'] = make_relative(old_child.end.xpath_end, root_xpath)
            yield cache

    def evaluate(self, context, base=None, allowed_styles=None):
        log.debug("Evaluating foreach %s %s", self.src, repr(self.start))
        content_elements = None
        last_paragraph = self.end.start.getparent().getnext()

        base = self.childs[0].field.start.getparent() if len(self.childs) > 0 else None
        if base is not None:
            root_xpath = self.childs[0].field.xpath_start.rsplit('/', 1)[0]
        else:
            root_xpath = None

        if len(self.content()) == 0:
            log.warning("Foreach with empty body: %s", self)
            return

        child_cache = list(self._build_cache(root_xpath))

        has_content = False
        for new_context in self.itervalues(context):
            has_content = True
            log.debug('Using context %r', new_context)

            if content_elements is None:
                log.debug("Working with original elements")
                content_elements = self.content()
                base = base if base is not None else content_elements[0]
                if not root_xpath:
                    root_xpath = base.getroottree().getpath(base)
                base_xpath = root_xpath
                log.debug('Root: %s (%s)', root_xpath, base)
                log.debug('Base: %s', base_xpath)
                content_elements = self.copy()
            else:
                log.debug("Working with cloned elements")
                for i, el1 in enumerate(content_elements):
                    el = deepcopy(el1)
                    last_paragraph.addprevious(el)
                    #last_paragraph.addnext(el)
                    #last_paragraph = el
                    if i == 0:
                        base = el
                        base_xpath = base.getroottree().getpath(base)
                log.debug('Base: %s', base_xpath)

            children = []
            for cache, old_child in zip(child_cache, self.childs):
                e_parent = cache['e_parent']
                s_parent = cache['s_parent']

                new_child = deepcopy(old_child)
                new_child.field._base = base

                # preserve xml tree connection (deepcopy makes a unlinked element)
                new_child.field._parent = cache['parent']

                new_child.field.xpath_start = make_abs(cache['field.xpath_start'], base_xpath)
                new_child.field.xpath_end = make_abs(cache['field.xpath_end'], base_xpath)

                if hasattr(new_child, 'start') and new_child.start != new_child.field:
                    log.debug('Update start for %s', new_child)
                    new_child.start._parent = s_parent
                    new_child.start._base = base

                    new_child.start.xpath_start = make_abs(cache['start.xpath_start'], base_xpath)
                    new_child.start.xpath_end = make_abs(cache['start.xpath_end'], base_xpath)

                if hasattr(new_child, 'end'):
                    log.debug('Update end for %s', new_child)
                    new_child.end._parent = e_parent
                    new_child.end._base = base

                    new_child.end.xpath_start = make_abs(cache['end.xpath_start'], base_xpath)
                    new_child.end.xpath_end = make_abs(cache['end.xpath_end'], base_xpath)

                children.append(new_child)

            for child in children:
                child.evaluate(new_context,
                               base=base,
                               allowed_styles=allowed_styles)
        if not has_content:
            self.remove_content()

        self.remove_fields()


def gen_tree(doc):
    fields = []
    for paragraph in doc.paragraphs:
        fields += paragraph.fields()

    for table in doc.tables:
        fields += table.fields()

    for textboxes in doc.textboxes:
        fields += textboxes.fields()

    fields = [f for f in fields if f.code == 'MERGEFIELD' and f.extra]

    container = Container()
    for f in fields:
        cmd = f.extra[0]
        cmd_type = cmd[0]
        cmd = cmd[1:]
        if cmd_type == '$':
            v = Variable(f, cmd)
            log.info("Variable found %s", v.path)
            container.childs.append(v)
        elif cmd_type == '#':
            if cmd.startswith('foreach'):
                m = re.search(r'^foreach\(\$?(.+)\s+in\s+\$(.+)\)$', cmd)

                container = ForEach(field=f, parent=container, src=m.group(2), dest=m.group(1))
                log.info("Foreach found %s in %s", container.dest, container.src)
                container.parent.childs.append(container)
            elif cmd.startswith('if'):
                m = re.search(r'^if\((.+)\)$', cmd)

                container = If(field=f, parent=container, src=m.group(1))
                log.info("If found %s", container.src)
                container.parent.childs.append(container)
            elif cmd.startswith('end'):
                container.end = f
                log.debug("End found %s (in %s)", f, unicode(container))
                log.debug("Has %d paragraphs", len(container.content()))
                container = container.parent

    if container.parent is not None:
        msg = "Missing #end for container %s" % container.start.default
        log.critical(msg)
        raise ParserException(msg, container)

    return container
