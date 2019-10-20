import pdfquery
import re
import os, shutil
import pyodbc
import datetime
import io
import re
from datetime import date
import smtplib
import openpyxl
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfpage import PDFPage
import random
from openpyxl import load_workbook




def extract_text_from_pdf(pdf_path):
    

    
    resource_manager = PDFResourceManager()
    fake_file_handle = io.StringIO()
    converter = TextConverter(resource_manager, fake_file_handle)
    page_interpreter = PDFPageInterpreter(resource_manager, converter)
    with open(pdf_path, 'rb') as fh:
        for page in PDFPage.get_pages(fh, 
                                      caching=True,
                                      check_extractable=True):
            page_interpreter.process_page(page)
        text = fake_file_handle.getvalue()
    converter.close()
    fake_file_handle.close()
    print(text)
    print('\n\n')
    
    if "Paramount Trading Corporation" in text:
        
        PO = re.search("PO(.*?)Invoice", text)
        PO = PO.group()
        PO = PO.replace("PO Ref : "," ")
        PO = PO.replace(" Invoice"," ")
        print(PO)
        
        date = re.search("Date :.{8}",text)
        date = date.group()
        date = date.replace("Date :"," ")
        print(date)
        
        name = "Paramount Trading Corporation"
        print(name)
        
        add = re.search("Billing Address (.*?)Date",text) 
        add = add.group().replace("Billing Address "," ").replace("Date"," ")
        print(add)
        
        inv = re.search("Invoice No(.*?)%", text)
        inv = inv.group().replace("Invoice No:- "," ").replace("%"," ")
        print(inv)
        
        #cpan = re.search("Customer PAN (.*?)Ship",text)
        #cpan = cpan.group().replace("Customer PAN No"," ").replace("Ship"," ")
        #print(cpan)
        
        #cgst = re.search("Customer GST (.*?)Customer", text)
        #cgst = cgst.group().replace("Customer GST No"," ").replace("Customer"," ")
        #print(cgst)
        
        
        gst = re.search("GST No : (.*?)PAN",text)
        gst = gst.group().replace("GST No : "," ").replace("PAN"," ")
        gst = gst.replace("Paramount Trading Corporation  "," ")
        print(gst)
        
        pan = re.search("PAN No : (.*?)Declaration",text)
        pan = pan.group().replace("PAN No : "," ").replace("Declaration"," ")
        print(pan)
        
        total = re.search("18%.{300}", text)
        total = total.group().split(".")
        total = total[1][2:] + "." + total[2][:2]
        print(total)
        
        tax = re.search("18%.{300}", text)
        tax = tax.group().split(".")
        tax[0] = tax[0].replace("18%", " ")
        tax = tax[0] + "." + tax[1][:2]
        print(tax)
        
        des = re.search("Paramount Trading Corporation(.*?)#8",text)
        des = des.group()
        des = des.replace("Description"," ")
        des = des.replace("#8"," ")
        des = des.replace("Commercial Invoice"," ")
        des = des.replace("Shipping Method"," ")
        des = des.replace("Mode of Payment"," ")
        des = des.replace("Shipment Date"," ")
        des = des.replace("Hero MotoCorp Ltd.C/o"," ")
        des = des.replace("The Grand New Delhi, Nelson Mandel Road, Vasant Kunj, Phase  IINew Delhi, India. Pin - 110070"," ")
        des = des.replace("Contact : Avinash +919557971063"," ")
        des = des.replace("Total"," ")
        des = des.replace("Paramount Trading Corporation"," ")
        des = des.replace("Road"," ")
        des = des.replace("11th June 2019"," ")
        des = des.replace("Hero MotoCorp Ltd."," ")
        des = des.replace("Customer PO Ref : "," ")
        des = des.replace(PO," ")
        des = des.replace("Invoice No:- "," ")
        des = des.replace("GST No : "," ")
        des = des.replace(gst," ")
        des = des.replace("PAN No : "," ")
        des = des.replace("%"," ")
        des = des.replace("Declaration:We declare that this invoice shows the actual price of the goodsdescribed and that all particulars are true and correct."," ")
        des = des.replace("Authorised Signatory"," ")
        des = des.replace("advance balance 60  ", " ")
        des = des.replace("against delivery", " ")
        des = des.replace(inv, " ")
        des = des.replace(pan, " ")
        des = des.replace("(round off)", " ")
        print(des)
        
    elif "SONATA" in text:
        
        PO = re.search("Cust PO Ref & Date(.*?)/", text)
        PO = PO.group().replace("Cust PO Ref & Date: "," ").replace("/"," ")
        print(PO)
        
        date = re.search("Invoice Date: (.*?)BILL",text)
        date = date.group().replace("Invoice Date: "," ").replace("BILL", " ")
        print(date)
        
        name = "SONATA INFORMATION TECHNOLOGY LIMITED"
        print(name)
        
        add = re.search("INVOICESONATA INFORMATION TECHNOLOGY LIMITED(.*?)TEL",text)
        add = add.group().replace("INVOICESONATA INFORMATION TECHNOLOGY LIMITED", " ").replace("TEL"," ")
        print(add)
        
        inv = re.search("Invoice No.:(.*?)Invoice",text)
        inv = inv.group().replace("Invoice No.:"," ").replace("Invoice"," ")
        print(inv)
        
        gst = re.search("GSTIN : (.*?)PAN",text)
        gst = gst.group().replace("GSTIN : "," ").replace("PAN"," ")
        print(gst)
        
        pan = re.search("Our PAN is (.*?)and",text)
        pan = pan.group().replace("Our PAN is "," ").replace("and"," ")
        print(pan)

        total = re.search("Total Invoice Value  (.*?)of",text)
        total = total.group().split(".")
        total[0] = total[0].replace("Total Invoice Value  "," ")
        total = total[0] + "." + total[1][:2]
        print(total)
        
        tax = re.search("Total Tax Value(.*?)Total",text)
        tax = tax.group().replace("Total Tax Value"," ").replace("Total", " ")
        print(tax)
        
        des = re.search("Description of Goods/Services(.*?)Each",text)
        des = des.group()
        des = des.replace("Description of Goods/Services", " ")
        des = des.replace("Each", " ")
        des = des.replace("Qty", " ")
        des = des.replace("UOM", " ")
        des = des.replace("Rate", " ")
        des = des.replace("(INR)", " ")
        des = des.replace("Amount", " ")
        print(des)
        
    elif "Concoct Human Resources Practitioners India" in text:
        
        PO = re.search("eWay Bill No#.{300}",text)
        PO = PO.group().split(" ")
        PO = PO[13]
        print(PO)
        
        date = re.search("eWay Bill No#.{300}",text)
        date = date.group().split(" ")
        date = date[12]
        print(date)
        
        name = "Concoct Human Resources Practitioners India"
        print(name)
        
        add = re.search("#(.*?)Proforma",text)
        add = add.group().replace("Proforma", " ")
        print(add)
        
        inv = re.search("Invoice No: (.*?)PAN",text)
        inv = inv.group().replace("Invoice No: "," ").replace("PAN"," ")
        print(inv)
        
        gst = re.search("IGST No#:(.*?)IEC",text)
        gst = gst.group().replace("IGST No#:", " ").replace("IEC"," ")
        print(gst)
        
        pan = re.search("PAN No: (.*?)GSTIN", text)
        pan = pan.group().replace("PAN No: ", " ").replace("GSTIN"," ")
        print(pan)
        
        total = re.search("Total Inc. of GST @ 18%(.*?)Amount",text)
        total = total.group().replace("Total Inc. of GST @ 18%"," ").replace("Amount"," ")
        print(total)
        
        tax = "Not given separately"
        print(tax)
        
        des = re.search("Particulars(.*?)Total",text)
        des = des.group()
        des = des.replace("Particulars"," ")
        des = des.replace("Product"," ")
        des = des.replace("S/N"," ")
        des = des.replace("No# of Units"," ")
        des = des.replace("Price Per Unit"," ")
        des = des.replace("GST @ 18%"," ")
        des = des.replace("Amount"," ")
        des = des.replace("(INR)"," ")
        des = des.split(".")
        #des = re.findall("[a-z]",des)
        l = len(des)
        for i in range(0,l-1):
            if "Unit" in des[i]:
                desi = des[i].split("Unit")
                desi = desi[0]
                print(desi)
        
    elif "MicroGenesis CADSoft" in text:
        
        PO = "Not given"
        print(PO)
        
        date = re.search("Despatched throughDated(.*?)Mode",text)
        date = date.group().replace("Despatched throughDated"," ").replace("Mode"," ")
        print(date)
        
        name = "MicroGenesis CADSoft"
        print(name)
        
        add = re.search("MicroGenesis CADSoft(.*?)MSMED",text)
        add = add.group().replace("MSMED"," ").replace("MicroGenesis CADSoft Pvt Ltd"," ")
        print(add)
        
        inv = re.search("Invoice No.(.*?)Delivery",text)
        inv = inv.group().replace("Invoice No.", " ").replace("Delivery", " ")
        print(inv)
        
        gst = re.search("GSTIN/UIN:(.*?)State",text)
        gst = gst.group().replace("GSTIN/UIN:", " ").replace("State"," ")
        print(gst)
        
        pan = re.search("Company's PAN :(.*?)Dec",text)
        pan = pan.group().replace("Company's PAN :", " ").replace("Dec"," ")
        print(pan)
        
        total = re.search("Total₹(.*?)No",text)
        total = total.group().replace("Total", " ").replace("No", " ")
        print(total)
        
        tax = re.search("IGST @ 18%(.*?)%",text)
        tax = tax.group().replace("IGST @ 18%"," ").replace("%"," ")
        print(tax)
        
        des = re.search("SACNo.Services(.*?)No", text)
        des = des.group().replace("SACNo.Services", " ").replace("No", " ")
        print(des)
        
        
        
    
        
        





extract_text_from_pdf(r'C:\Users\garv2\Desktop\MicroGenesis CADSoft Pvt Ltd\DL_0046_19-20.pdf')   
