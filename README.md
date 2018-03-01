batchpdfconvert
===============

Copyright 2017 Pontus Lurcock. Released under the MIT License;
see the file LICENSE for details.

Rationale
---------

I sometimes come across PowerPoint files containing useful information
that I'd like to keep, but PowerPoint is not a good archival format. In
order to ensure future readability I prefer to convert these files to
PDF. The conversion can be done manually be opening a PowerPoint file in
LibreOffice and exporting it as a PDF, but this becomes inconvenient
when many files need to be converted; there are also some minor
tweaks I like to perform before the conversion, which would be tedious
to do by hand.

`batchpdfconvert` provides a command-line interface to the LibreOffice
conversion functionality via the UNO API. It makes the following
modifications to the files when saving:

- Many PowerPoint files are unnecessarily large, because they include
  excessively high-resolution bitmap graphics. Usually the text is the
  content I'm mainly interested in, so `batchpdfconvert` exports
  all images to JPEGs with an 80% compression ratio.

- LibreOffice sometimes converts text with the "shadow" decoration to
  two separate, overlapping, selectable text objects. Some PDF readers
  make it impossible to select just one of these text lines, so any
  text copied out of the PDF has each word, or each character,
  duplicated. `batchpdfconvert` therefore removes the shadow effect
  throughout the file before performing the conversion.

- Occasionally I find an object in a PowerPoint file with a negative
  x position, so that its leftmost portion is clipped off the edge of
  the slide. `batchpdfconvert` finds such objects and changes the x
  position to 0, moving them onto the page.

Installation
------------

batchpdfconvert is developed in Java as a Netbeans project with a build
file in Ant. Netbeans is not required to build the project.
batchpdfconvert depends on various LibreOffice jars, which are provided
by the `libreoffice-java-common` and `ure` packages in Ubuntu.

    sudo apt install openjdk-8-jdk ure libreoffice-java-common ant
    ant fatjar

Package `ure` also provides `/usr/lib/libreoffice/program/libjpipe.so`.

Usage
-----

    java -Djava.library.path=/usr/lib/libreoffice/program/ -jar batchpdfconvert.jar input.ppt output.pdf
