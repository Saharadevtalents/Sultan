# -*- coding: utf-8 -*-
import xlwt
import io
from io import StringIO
from odoo import api, exceptions, fields, models, _
from tempfile import TemporaryFile
import base64
from xlrd import open_workbook
import logging
import xlrd
import os
import re
import time
from odoo.addons import decimal_precision as dp

_logger = logging.getLogger(__name__)


class product(models.Model):
    _inherit = "product.product"
    
    list_price = fields.Float(
        'Sales Price', default=1.0,
        digits=dp.get_precision('Product Price'),
        help="Base price to compute the customer price. Sometimes called the catalog price.")


class ExportFile(models.TransientModel):
    _name = "export.excel.report"
    
    data = fields.Binary('File')
    name= fields.Char("Name")

class UpdateProducts(models.TransientModel):
    _name = "product.update"
    
    name = fields.Char('File Name')
    description = fields.Text('Description')
    file = fields.Binary('File')
    
    def get_product_barcode(self, barcode):
        product_obj = self.env['product.product']
        product_ids = product_obj.search([('barcode','=',barcode)])
        return product_ids and product_ids[0] or False
    
    def update_product(self, product, list_price, standard_price, commission):
        product.write({'list_price': list_price, 
                       'standard_price': standard_price,
                       'commission_amount': commission})
        return True

    @api.multi
    def update_list_price(self):
        """ Open File list price """
        file_data = base64.decodestring(self.file)
        book = open_workbook(file_contents=file_data)
        xl_sheet = book.sheet_by_index(0)
        remaining = xl_sheet.nrows
        
        lines_notexist = []        
        for rx in range(1, xl_sheet.nrows):
            remaining -= 1
            barcode = xl_sheet.cell_value(rx, 1)
            product = self.get_product_barcode(str(barcode))
            if product:
                price = xl_sheet.cell_value(rx, 4)
                cost = xl_sheet.cell_value(rx, 3)
                commission = xl_sheet.cell_value(rx, 5)
                self.update_product(product, price, cost, commission)
            else:
                line = {
                    'ItemNo': xl_sheet.cell_value(rx, 0),
                    'Barcode': xl_sheet.cell_value(rx, 1),
                    'Qty': xl_sheet.cell_value(rx, 2),
                    'Avr_Cost': xl_sheet.cell_value(rx, 3),
                    'SalesPrice': xl_sheet.cell_value(rx, 4),
                    'Reward Value': xl_sheet.cell_value(rx, 5),  
                }
                lines_notexist.append(line)
            logging.info('%s rows remaining' % remaining)
            
        """ generate File not exist barcode """
        return self.generate_excel_file(lines_notexist)
            
    def generate_excel_file(self, lines_notexist):
        if self._context is None:
            self._context = {}
        workbook = xlwt.Workbook(encoding='utf-8')
        sheet = workbook.add_sheet('Sheet_1')
        # write  the header of report
        row=self.write_header(sheet)
        #write body
        
        self.write_body(sheet,row, lines_notexist)
        
        file_data = io.BytesIO()
        o = workbook.save(file_data)
        out = base64.encodestring(file_data.getvalue())
        
        name = 'Products -NoBarcode '+ str(time.strftime("%Y-%m-%d %H:%M:%S")) + '.xls'
        
        this = self.env['export.excel.report'].create({'data': out, 'name': name })
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'export.excel.report',
            'view_mode': 'form',
            'view_type': 'form',
            'res_id': this.id,
            'views': [(False, 'form')],
            'target': 'new',
        }
        
    def write_header(self,sheet):
        row = 0
        sheet.write(row,0,'ItemNo')
        sheet.write(row,1,'Barcode')
        sheet.write(row,2,'Qty') 
        sheet.write(row,3,'Avr_Cost')
        sheet.write(row,4,'SalesPrice')
        sheet.write(row,5,'Reward Value')
        return row
    
    def write_body(self,sheet,row, lines):
        for line in lines:
            row += 1
            if line['ItemNo']:
                sheet.write(row,0,int(line['ItemNo']))
            sheet.write(row,1,str(line['Barcode']))
            sheet.write(row,2,line['Qty'])
            sheet.write(row,3,line['Avr_Cost'])
            sheet.write(row,4,line['SalesPrice'],)
            sheet.write(row,5,line['Reward Value'])
        return True


    def get_product_code(self, code):
        product_obj = self.env['product.product']
        product_ids = product_obj.search([('default_code','like',code)])
        return product_ids and product_ids[0] or False
    
    def update_product_barcode(self, product, barcode):
        product.write({'barcode': barcode})
        return True
    
    @api.multi
    def update_barcode_product(self):
        """ Open File barcode """
        file_data = base64.decodestring(self.file)
        book = open_workbook(file_contents=file_data)
        xl_sheet = book.sheet_by_index(0)
        remaining = xl_sheet.nrows
        
        for rx in range(1, xl_sheet.nrows):
            remaining -= 1
            default_code = xl_sheet.cell_value(rx, 0)
            product = self.get_product_code(int(default_code))
            if product:
                barcode = xl_sheet.cell_value(rx, 3)
                self.update_product_barcode(product, barcode)
            else:
                continue
            logging.info('%s rows remaining' % remaining)
        
        return  True
                
