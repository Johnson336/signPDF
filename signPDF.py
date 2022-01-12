# -*- coding: utf-8 -*-
""" Python PDF Manipulator test
"""
from PyPDF4 import PdfFileWriter, PdfFileReader
import io
import os
import sys
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import date
import locale
import ghostscript
import win32print
import win32api
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd


# Return printer name to use. If -printer is set it will return this value only if value match with available
# printers list. Return a error if -printer not in list. If no printer specified, retrieve default printer and return
# its name. Sometime default printer is on USELESS_PRINTER list so first printer return by getAvailablePrinters() is
# return. If no printer is return display an error.
def getPrinter():
    default_printer = win32print.GetDefaultPrinter()
    return default_printer

# Use GhostScript API to silent print .pdf and .ps. Use win32api to print .txt. Return a error if printing failed or
# file ext doesn't match.
def printFile(filepath):
    try:
        if os.path.splitext(filepath)[1] in [".pdf", ".ps"]:
            args = [
                "-dPrinted", "-dBATCH", "-dNOSAFER", "-dNOPAUSE", "-dNOPROMPT"
                                                                  "-q",
                "-dNumCopies#1",
                "-sDEVICE#mswinpr2",
                f'-sOutputFile#"%printer%{getPrinter()}"',
                f'"{filepath}"'
            ]

            encoding = locale.getpreferredencoding()
            args = [a.encode(encoding) for a in args]
            ghostscript.Ghostscript(*args)
        elif os.path.splitext(filepath)[1] in [".txt"]:
            # '"%s"' % enable to encapsulate string with quote
            win32api.ShellExecute(0, "printto", '"%s"' % filepath, '"%s"' % getPrinter(), ".", 0)
        return True

    except:
        print("Printing error for file: ", '"%s"' % filepath, "| Printer: ", '"%s"' % getPrinter())
        return False

""" Tkinter dialog window functions """

root = tk.Tk()
root.title("Decrypt and sign voucher wizard")
root.resizable(False, False)
root.geometry('300x100')
#text = tk.Text(root, height=12)
#text.grid(column=0, row=0, sticky='nsew')

def open_file():
    filetypes = (('PDF Files', '*.pdf'),('All Files', '*.*'))
    filepath = fd.askopenfilename(filetypes=filetypes)
    #text.insert('1.0', filepath.readlines())
    print("Decrypting "+filepath+"...")
    #text.pack()
    
    
    if not os.path.isfile(filepath) and not os.path.exists(filepath):
        print("Path provided is not a file path.")
        sys.exit(2)
    if not printFile(filepath):
        sys.exit(1)
        
def sign_file():
    filetypes = (('PDF Files', '*.pdf'),('All Files', '*.*'))
    filepath = fd.askopenfilename(filetypes=filetypes)
    print("Signing "+filepath+"...")
    #text.pack()
    
    initial_sig = io.BytesIO()
    can = canvas.Canvas(initial_sig, pagesize=letter)
    can.drawString(180, 364, "KN")
    can.save()
    
    signed_sig = io.BytesIO()
    can2 = canvas.Canvas(signed_sig, pagesize=letter)
    can2.drawString(180, 561, "Kim Niermeyer")
    can2.drawString(180, 528, "vouchers@wgu.edu")
    today = date.today()
    can2.drawString(420, 494, today.strftime("%m/%d/%Y"))
    pdfmetrics.registerFont(TTFont('Edwardian', 'C:\\WINDOWS\\FONTS\\ITCEDSCR.TTF'))
    can2.setFont('Edwardian', 20)
    can2.drawString(200, 494, "Kim Niermeyer")
    can2.save()
    signed_sig.seek(0)
    new_pdf2 = PdfFileReader(signed_sig)
    
    #move to the beginning of the StringIO buffer
    initial_sig.seek(0)
    
    # create a new PDF with Reportlab
    new_pdf = PdfFileReader(initial_sig)
    # read your existing PDF
    try:
        existing_pdf = PdfFileReader(open(filepath, "rb"),strict=False)
    except:
        print("Error opening existing file.")
        return False
    output = PdfFileWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    output.addPage(page)
    page = existing_pdf.getPage(1)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
    page = existing_pdf.getPage(2)
    page.mergePage(new_pdf2.getPage(0))
    output.addPage(page)
    # finally, write "output" to a real file
    try:
        filetypes = (('PDF Files', '*.pdf'),('All Files', '*.*'))
        filepath = fd.asksaveasfilename(filetypes=filetypes)
        outputStream = open(filepath, "wb")
    except:
        print("Error signing file.")
        return False
    output.write(outputStream)
    outputStream.close()
    print("File signed, operation completed.")

decrypt_button = ttk.Button(
    root,
    text='Choose file to decrypt',
    command=open_file
    )
decrypt_button.grid(column=0, row=1, sticky='ws', padx=10, pady=10)
sign_button = ttk.Button(
    root,
    text='Choose file to sign',
    command=sign_file
    )
sign_button.grid(column=0, row=2, sticky='ws', padx=10, pady=10)
root.mainloop()


