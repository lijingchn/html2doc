#!/usr/bin/env python
# encoding: utf-8

import pdfkit

#pdfkit.from_url("http://stackoverflow.com/questions/23359083/how-to-convert-webpage-into-pdf-by-using-python", "ics.pdf")
pdfkit.from_file("word_pdf.html", "word_pdf.pdf")

