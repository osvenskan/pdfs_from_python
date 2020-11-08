#!/usr/bin/env python3
"""This is a demo of how to manipulate the content of a LibreOffice/OpenOffice document using pure
Python to alter the document's XML. Optionally, the document can be rendered to PDF. This is
companion code to a series of blog posts that I wrote:
http://blog.pyspoken.com/2016/07/27/creating-pdf-documents-using-libreoffice-and-python/
http://blog.pyspoken.com/2016/08/12/creating-pdf-documents-using-libreoffice-and-python-part-2/
http://blog.pyspoken.com/2016/10/07/creating-pdf-documents-using-libreoffice-and-python-part-3/

I also gave a talk on this topic at PyOhio 2016:
https://www.youtube.com/watch?v=uwWvz5QLtiI

This code was written under Python 3.5.

The code below uses ElementTree to manipulate XML. I used ElementTree because it's part of the
standard library and I want people to be able to run this code without installing any 3rd party
packages. For more robust code, I would recommend lxml instead of ElementTree because it offers
some features (like the ability to navigate to a node's parents and siblings) that would make
the code below simpler.

This code is released under a BSD license; see accompanying LICENSE file for details.
"""

import tempfile
import zipfile
import os
import xml.etree.ElementTree as ET
import subprocess

# PATH_TO_EXECUTABLE is the path to the LibreOffice executable. If set, the code will launch
# LibreOffice headlessly to convert to PDF the document that this process creates.
PATH_TO_EXECUTABLE = None
# PATH_TO_EXECUTABLE = '/Applications/LibreOffice.app/Contents/MacOS/soffice'   # macOS
# PATH_TO_EXECUTABLE = '/usr/bin/soffice'                                       # Some Linuxes
# PATH_TO_EXECUTABLE = 'C:/Program Files/Office/program/soffice'                # Some Windowses

# These are the namespaces used in the XML of this simple document.
NAMESPACES = {'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
              'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
              'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
              'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
              }
# It's not super clear from the documentation, but register_namespace() only affects serialization
# of XML. When calling e.g. find(), we still have to pass NAMESPACES explicitly.
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


class Helper():
    """A class to provide some utility features for manipulating the XML."""

    def __init__(self, path):
        self.tree = ET.parse(os.path.join(path, 'content.xml'))

        self.root = self.tree.getroot()

        # We need to be able to find the parent of a node, and ElementTree doesn't directly support
        # that. Instead I'll build a mapping here using a solution stolen^W borrowed from
        # watbywbarif's stackoverflow post: http://stackoverflow.com/a/12533735
        # This technique has limitations, because the map is static and doesn't reflect changes I
        # make to the tree, but it's good enough for this demo. The 3rd party package lxml has
        # full support for navitgating from children to parents.
        self.parent_map = dict((child, parent) for parent in self.tree.getiterator()
                               for child in parent)

    def populate_bookmark(self, name, contents):
        """Given a LibreOffice bookmark name, populates it with the contents param.

        This operates on bookmarks that already have content; they have a bookmark-start and
        a bookmark-end tag like the example below. One can also define a LibreOffice bookmark with
        no content. Writing code to populate those bookmarks is left as an exercise for the reader.

        Example of a bookmark with content:
            <text:bookmark-start text:name="fox_type_placeholder" />
            <text:s />
            <text:bookmark-end text:name="fox_type_placeholder" />

        Example of a bookmark with no content:
            <text:bookmark text:name="my_bookmark" />
        """
        bookmark_start = self.root.find('.//text:bookmark-start[@text:name="{}"]'.format(name),
                                        NAMESPACES)

        # Find the next sibling.
        bookmark_parent = self.parent_map[bookmark_start]

        for i, child in enumerate(bookmark_parent):
            if child is bookmark_start:
                break

        next_sibling = bookmark_parent[i + 1]
        # Change it to a span and populate it.
        next_sibling.tag = 'text:span'
        next_sibling.text = contents


# =================================    Main starts here   =================================


with tempfile.TemporaryDirectory() as tempdir:
    # Unzip the odt. This is equivalent of this command line:
    #     unzip input.odt -d [tempdir]
    with zipfile.ZipFile('input.odt') as the_odt:
        the_odt.extractall(tempdir)
        odt_filenames = the_odt.namelist()

    helper = Helper(tempdir)

    # Add text to the two bookmarks that I created while editing the doc in LibreOffice.
    helper.populate_bookmark('fox_type_placeholder', 'quick brown')
    helper.populate_bookmark('dog_type_placeholder', 'lazy')

    # Add a paragraph. Start by finding either one of my bookmarks.
    bookmark = helper.root.find('.//text:bookmark-start', NAMESPACES)
    # The bookmark parent is a paragraph. Clone it.
    paragraph = helper.parent_map[bookmark]
    new_paragraph = ET.Element('text:p', paragraph.attrib)
    new_paragraph.text = 'No one expects the Spanish Inquisition!'
    parent = helper.parent_map[paragraph]
    for i, child in enumerate(parent):
        if child is paragraph:
            break

    parent.insert(i + 1, new_paragraph)

    # OK, content.xml has been altered by the added paragraph and the calls to populate_bookmark().
    # Write it back to disk.
    with open(os.path.join(tempdir, 'content.xml'), 'wb') as f:
        helper.tree.write(f, 'utf-8', True)

    # Zip the doc back into a single file. A command line almost-equivalent is --
    #    zip -r target_filename.odt *
    # That command doesn't respect the subtle mimetype rule described below, so it's not quite
    # correct. It's usually OK if you're just testing.

    # Per the OpenDocument standard, "The 'mimetype' file shall be the first file of the zip file.
    # It shall not be compressed, and it shall not use an 'extra field' in its header."
    # ref: Open Document Format for Office Applications (OpenDocument) Version 1.2, section 3.3
    with zipfile.ZipFile('output.odt', 'w') as output_file:
        # Write the mimetype file
        output_file.write(os.path.join(tempdir, 'mimetype'), 'mimetype')

        # Write everything else compressed.
        for dirpath, dirnames, filenames in os.walk(tempdir):
            for filename in filenames:
                if filename != 'mimetype':
                    absolute_filename = os.path.join(dirpath, filename)
                    # The zip filename must be relative to the root of the .zip, e.g. mimetype
                    # is in the root, image files are in Pictures/foo.png, etc.
                    zip_filename = os.path.relpath(absolute_filename, tempdir)
                    output_file.write(absolute_filename, zip_filename, zipfile.ZIP_DEFLATED)

if PATH_TO_EXECUTABLE:
    final_filename = 'output.pdf'
    # Launch LibreOffice to convert the doc to PDF. The 3rd party unoconv package makes this
    # a bit easier.
    # Note: This step will fail silently if LibreOffice is already running.
    subprocess.call((PATH_TO_EXECUTABLE, "--headless", "--convert-to", "pdf:writer_pdf_Export",
                    "--outdir", os.getcwd(), 'output.odt'))
else:
    final_filename = 'output.odt'

print('Done! Output file is in {}.'.format(final_filename))
