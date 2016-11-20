#!/usr/bin/env python
# encoding: utf-8

import weasyprint
weasyprint.HTML("http://kaito-kidd.com/2015/03/12/python-html2pdf/").write_pdf("test1.pdf")
weasyprint.HTML("http://kaito-kidd.com/2015/03/12/python-html2pdf/").write_png("test1.png")
