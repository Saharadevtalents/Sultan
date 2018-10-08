# -*- coding: utf-8 -*-

from odoo import api, exceptions, fields, models, _

import logging
import xlrd
import os
import re
import ast

_logger = logging.getLogger(__name__)


class Importer(models.TransientModel):
    _name = "product.import"


    # @api.model
    # def _create_product(self, name,brand_name, values={}):
    #     product_obj = self.env["product.template"]
    #     brand_obj = self.env["product.brand"]
    #     brand_id=brand_obj.search([('name', '=', brand_name)], limit=1)
    #     if name:
    #         if brand_name:
    #             record = product_obj.search([('name', '=', name),('product_brand_id','=',brand_id.id)], limit=1)
    #         else:
    #             record = product_obj.search([('name', '=', name)], limit=1)
    #     else:
    #         return product_obj
    #
    #     if not record:
    #         default_values = product_obj.default_get(list(product_obj._fields))
    #         values.update({'name': name})
    #         values.update({'product_brand_id': brand_id.id})
    #         default_values.update(values)
    #         record = product_obj.create(default_values)
    #     return record

    @api.model
    def _create_product(self, name, values={}):
        obj = self.env["product.template"]
        if not name:
            return obj
        record = False
        if not record:
            default_values = obj.default_get(list(obj._fields))
            values.update({'name': name})
            default_values.update(values)
            record = obj.create(default_values)
        return record

    @api.model
    def _create_product_product(self, name, values={}):
        obj = self.env["product.product"]
        if not name:
            return obj
        record = False
        if not record:
            default_values = obj.default_get(list(obj._fields))
            values.update({'name': name})
            default_values.update(values)
            record = obj.create(default_values)
        return record

    @api.model
    def _create_obj(self, obj_name, name, values={}):
        obj = self.env[obj_name]
        if not name:
            return obj
        record = obj.search([('name', '=', name)], limit=1)
        if not record:
            default_values = obj.default_get(list(obj._fields))
            values.update({'name': name})
            default_values.update(values)
            record = obj.create(default_values)
        return record

    @api.model
    def _add_variant(self, product_template, attribute_id, attribute_value):
        attribute_value_id = self.env['product.attribute.value'].search(
            [('attribute_id', '=', attribute_id), ('name', '=', attribute_value)], limit=1)
        if not attribute_value_id:
            attribute_value_id = self._create_obj('product.attribute.value', attribute_value,
                                                  {'attribute_id': attribute_id})
        attribute_added = False
        att = False
        for attribute_line in product_template.attribute_line_ids:
            if attribute_line.attribute_id.id == attribute_id:
                attribute_line.value_ids |= attribute_value_id
                if attribute_value_id.id in attribute_line.value_ids.ids:
                    attribute_added = True
                    break
        if not attribute_added:
            product_template.attribute_line_ids.create({
                'product_tmpl_id': product_template.id,
                'attribute_id': attribute_id,
                'value_ids': [(4, attribute_value_id.id)],
            })
            product_template.with_context(create_from_tmpl=True).create_variant_ids()
        product_template.create_variant_ids()
    @api.model
    def _find_product(self, domain):
        product_obj = self.env['product.template']
        products = product_obj.search(domain)

        return products
        # for product in products:
        #     if set(product.attribute_value_ids.ids) == set(attributes):
        #         return product

    # @api.multi
    # def import_product_sultan(self):
    #     product_template = self.env['product.template'].search([('id', '=', 1939)], limit=1)
    #     attribute_value_id = self.env['product.attribute.value'].search([('name', '=', u"شتوية")], limit=1)
    #     attribute_id = self.env['product.attribute'].search([('name', '=', u"الموسم")], limit=1)
    #     product_template.attribute_line_ids.create({
    #         'product_tmpl_id': product_template.id,
    #         'attribute_id': attribute_id.id,
    #         'value_ids': [(4, attribute_value_id.id)],
    #     })

    def get_internal_category(self,string):
        Liste=string.split("\\")
        product_category_obj = self.env["product.category"]
        j=0
        ch=""
        for l in Liste:
            if ch=="":
                category=product_category_obj.search([('complete_name', '=', l)], limit=1)
                ch= l
            else:
                ch= '%s / %s' % (ch, l)
                category=product_category_obj.search([('complete_name', '=', ch)], limit=1)
            if not category:
               if j==0:
                  category=product_category_obj.create({'name':l})
               else:
                  category=product_category_obj.create({'name':l,'parent_id':parent_id})
               parent_id=category.id
            else:
                parent_id=category.id
            j+=1

        return category.id

    # def get_pos_category(self,string):
    #     product_tmpl_obj = self.env['product.template']
    #     pos_category_obj = self.env['pos.category']
    #     p=pos_category_obj.search([('id', '=', 2)]).name_get()
    #     # string=u"السلطان\الاثواب"
    #     product_tmpl_obj.search([('id', '=', 2188)], limit=1).write({'description_sale':p[0][1]})
    #
    #     Liste=string.split("\\")
    #     product_category_obj = self.env["pos.category"]
    #     j=0
    #     ch=""
    #     for l in Liste:
    #         if ch=="":
    #             category=product_category_obj.search([('complete_name', '=', l)], limit=1)
    #             ch= l
    #         else:
    #             ch= '%s / %s' % (ch, l)
    #             category=product_category_obj.search([('complete_name', '=', ch)], limit=1)
    #         if not category:
    #            if j==0:
    #               category=product_category_obj.create({'name':l})
    #            else:
    #               category=product_category_obj.create({'name':l,'parent_id':parent_id})
    #            parent_id=category.id
    #         else:
    #             parent_id=category.id
    #         j+=1
    #
    #     return category.id


    @api.multi
    def import_product_attribute(self):
        self._create_obj('product.attribute', u"المقاس", values={})
        self._create_obj('product.attribute', u"اللون", values={})
        self._create_obj('product.attribute', u"السروال", values={})
        self._create_obj('product.attribute', u"الموديل", values={})
        self._create_obj('product.attribute', u"الموسم", values={})
        self._create_obj('product.attribute', u"نوعية الشغل", values={})
        self._create_obj('product.attribute', u"نوعية القماش", values={})


    @api.multi
    def import_product_sultan_badla_arabe(self):
        book = xlrd.open_workbook(os.path.dirname(__file__) + "/Sultan Products.xlsx")
        xl_sheet = book.sheet_by_index(0)
        remaining = xl_sheet.nrows

        for rx in range(1, xl_sheet.nrows):
            # if row has no product skip
            remaining -= 1
            if not xl_sheet.cell_value(rx, 3):
                continue
            name = xl_sheet.cell_value(rx, 3)
            domain=[('name', '=', name)]
            if len(xl_sheet.cell_value(rx, 0))!=0 :
                brand_obj = self.env["product.brand"]
                brand = brand_obj.search([('name', '=', xl_sheet.cell_value(rx, 0))], limit=1)
                if not brand:
                    brand_id = self._create_obj('product.brand', xl_sheet.cell_value(rx, 0)).id
                else:
                    brand_id = brand.id

                domain+=[(('product_brand_id', '=', brand_id))]
            internal_category=self.get_internal_category(str(xl_sheet.cell_value(rx, 7)))
            pos_category=self._create_obj('pos.category',str(xl_sheet.cell_value(rx, 8))).id


            product_template = self._find_product(domain)
            if product_template:

                if len(xl_sheet.cell_value(rx, 1)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"الموسم")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 1))

                if len(xl_sheet.cell_value(rx, 2)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"الموديل")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 2))

                if len(xl_sheet.cell_value(rx, 4)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"السروال")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 4))

                if len(xl_sheet.cell_value(rx, 5)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"اللون")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 5))

                if xl_sheet.cell_value(rx, 6):
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"المقاس")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 6)) )


            if not product_template:
                values = {}

                if brand_id:
                    values['product_brand_id'] = brand_id

                if internal_category:
                    values['categ_id'] = internal_category

                if pos_category:
                    values['pos_categ_id'] = pos_category

                product_template=self._create_product(name, values=values)
                if len(xl_sheet.cell_value(rx, 1)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"الموسم")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 1))

                if len(xl_sheet.cell_value(rx, 2)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"الموديل")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 2))

                if len(xl_sheet.cell_value(rx, 4)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"السروال")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 4))

                if len(xl_sheet.cell_value(rx, 5)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"اللون")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 5))

                if xl_sheet.cell_value(rx, 6):
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"المقاس")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 6)) )

                product_template.generate_barcode()

            logging.info('%s rows remaining' % remaining)

    @api.multi
    def import_product_sultan_thawb(self):
        book = xlrd.open_workbook(os.path.dirname(__file__) + "/Sultan Products.xlsx")
        xl_sheet = book.sheet_by_index(1)
        remaining = xl_sheet.nrows


        for rx in range(1, xl_sheet.nrows):
            # if row has no product skip
            remaining -= 1
            if not xl_sheet.cell_value(rx, 3):
                continue
            name = xl_sheet.cell_value(rx, 3)

            brand_obj = self.env["product.brand"]
            brand = brand_obj.search([('name', '=', xl_sheet.cell_value(rx, 0))], limit=1)
            if not brand:
                brand_id = self._create_obj('product.brand', xl_sheet.cell_value(rx, 0)).id
            else:
                brand_id = brand.id

            internal_category=self.get_internal_category(str(xl_sheet.cell_value(rx, 7)))
            pos_category=self._create_obj('pos.category',str(xl_sheet.cell_value(rx, 8))).id


            product_template = self._find_product([('name', '=', name), ('product_brand_id', '=', brand_id)])
            if product_template:
                if len(xl_sheet.cell_value(rx, 1)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"الموسم")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 1))

                if len(xl_sheet.cell_value(rx, 2)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"الموديل")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 2))

                if len(xl_sheet.cell_value(rx, 4)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"السروال")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 4))

                if len(xl_sheet.cell_value(rx, 5)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"اللون")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 5))

                if xl_sheet.cell_value(rx, 6):
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"المقاس")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 6)) )


            if not product_template:
                values = {}

                if brand_id:
                    values['product_brand_id'] = brand_id

                if internal_category:
                    values['categ_id'] = internal_category

                if pos_category:
                    values['pos_categ_id'] = pos_category

                product_template=self._create_product(name, values=values)
                if len(xl_sheet.cell_value(rx, 1)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"الموسم")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 1))

                if len(xl_sheet.cell_value(rx, 2)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"الموديل")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 2))

                if len(xl_sheet.cell_value(rx, 4)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"السروال")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 4))

                if len(xl_sheet.cell_value(rx, 5)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"اللون")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 5))

                if xl_sheet.cell_value(rx, 6):
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"المقاس")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 6)) )

                product_template.generate_barcode()

            logging.info('%s rows remaining' % remaining)


    @api.multi
    def import_product_sultan_horka(self):
        book = xlrd.open_workbook(os.path.dirname(__file__) + "/Sultan Products.xlsx")
        xl_sheet = book.sheet_by_index(2)
        remaining = xl_sheet.nrows


        for rx in range(1, xl_sheet.nrows):
            # if row has no product skip
            remaining -= 1
            if not xl_sheet.cell_value(rx, 2):
                continue
            name = xl_sheet.cell_value(rx, 2)

            brand_obj = self.env["product.brand"]
            brand = brand_obj.search([('name', '=', xl_sheet.cell_value(rx, 0))], limit=1)
            if not brand:
                brand_id = self._create_obj('product.brand', xl_sheet.cell_value(rx, 0)).id
            else:
                brand_id = brand.id

            internal_category=self.get_internal_category(str(xl_sheet.cell_value(rx, 6)))
            #pos_category=self._create_obj('pos.category',str(xl_sheet.cell_value(rx, 8))).id


            product_template = self._find_product([('name', '=', name), ('product_brand_id', '=', brand_id)])
            if product_template:
                # if len(xl_sheet.cell_value(rx, 1)) != 0:
                #     attribute_id = self.env['product.attribute'].search([('name', '=', u"الموسم")], limit=1).id
                #     self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 1))

                if len(str(xl_sheet.cell_value(rx, 1))) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"الموديل")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 1)))

                if len(xl_sheet.cell_value(rx, 3)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"السروال")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 3))

                if len(xl_sheet.cell_value(rx, 4)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"اللون")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 4))

                if xl_sheet.cell_value(rx, 5):
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"المقاس")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 5)) )


            if not product_template:
                values = {}

                if brand_id:
                    values['product_brand_id'] = brand_id

                if internal_category:
                    values['categ_id'] = internal_category

                # if pos_category:
                #     values['pos_categ_id'] = pos_category

                product_template=self._create_product(name, values=values)
                # if len(xl_sheet.cell_value(rx, 1)) != 0:
                #     attribute_id = self.env['product.attribute'].search([('name', '=', u"الموسم")], limit=1).id
                #     self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 1))

                if len(str(xl_sheet.cell_value(rx, 1))) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"الموديل")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 1))

                if len(xl_sheet.cell_value(rx, 3)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"السروال")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 3))

                if len(xl_sheet.cell_value(rx, 4)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"اللون")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 4))

                if xl_sheet.cell_value(rx, 5):
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"المقاس")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 5)) )

                product_template.generate_barcode()

            logging.info('%s rows remaining' % remaining)

    @api.multi
    def import_product_sultan_dakhilia(self):
        book = xlrd.open_workbook(os.path.dirname(__file__) + "/Sultan Products.xlsx")
        xl_sheet = book.sheet_by_index(3)
        remaining = xl_sheet.nrows


        for rx in range(1, xl_sheet.nrows):
            # if row has no product skip
            remaining -= 1
            if not xl_sheet.cell_value(rx, 1):
                continue
            name = xl_sheet.cell_value(rx, 1)

            brand_obj = self.env["product.brand"]
            brand = brand_obj.search([('name', '=', xl_sheet.cell_value(rx, 0))], limit=1)
            if not brand:
                brand_id = self._create_obj('product.brand', xl_sheet.cell_value(rx, 0)).id
            else:
                brand_id = brand.id

            internal_category=self.get_internal_category(str(xl_sheet.cell_value(rx, 4)))
            #pos_category=self._create_obj('pos.category',str(xl_sheet.cell_value(rx, 8))).id


            product_template = self._find_product([('name', '=', name), ('product_brand_id', '=', brand_id)])
            if product_template:
                # if len(xl_sheet.cell_value(rx, 1)) != 0:
                #     attribute_id = self.env['product.attribute'].search([('name', '=', u"الموسم")], limit=1).id
                #     self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 1))

                # if len(str(xl_sheet.cell_value(rx, 1))) != 0:
                #     attribute_id = self.env['product.attribute'].search([('name', '=', u"الموديل")], limit=1).id
                #     self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 1)))

                # if len(xl_sheet.cell_value(rx, 3)) != 0:
                #     attribute_id = self.env['product.attribute'].search([('name', '=', u"السروال")], limit=1).id
                #     self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 3))

                if len(xl_sheet.cell_value(rx, 2)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"اللون")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 2))

                if xl_sheet.cell_value(rx, 3):
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"المقاس")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 3)) )


            if not product_template:
                values = {}

                if brand_id:
                    values['product_brand_id'] = brand_id

                if internal_category:
                    values['categ_id'] = internal_category

                # if pos_category:
                #     values['pos_categ_id'] = pos_category

                product_template=self._create_product(name, values=values)
                # if len(xl_sheet.cell_value(rx, 1)) != 0:
                #     attribute_id = self.env['product.attribute'].search([('name', '=', u"الموسم")], limit=1).id
                #     self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 1))

                # if len(str(xl_sheet.cell_value(rx, 1))) != 0:
                #     attribute_id = self.env['product.attribute'].search([('name', '=', u"الموديل")], limit=1).id
                #     self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 1)))

                # if len(xl_sheet.cell_value(rx, 3)) != 0:
                #     attribute_id = self.env['product.attribute'].search([('name', '=', u"السروال")], limit=1).id
                #     self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 3))

                if len(xl_sheet.cell_value(rx, 2)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"اللون")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 2))

                if xl_sheet.cell_value(rx, 3):
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"المقاس")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 3)) )

                product_template.generate_barcode()

            logging.info('%s rows remaining' % remaining)

    @api.multi
    def import_product_sultan_accessoires(self):
        book = xlrd.open_workbook(os.path.dirname(__file__) + "/Sultan Products.xlsx")
        xl_sheet = book.sheet_by_index(4)
        remaining = xl_sheet.nrows


        for rx in range(1, xl_sheet.nrows):
            # if row has no product skip
            remaining -= 1
            if not xl_sheet.cell_value(rx, 1):
                continue
            name = xl_sheet.cell_value(rx, 1)

            # brand_obj = self.env["product.brand"]
            # brand = brand_obj.search([('name', '=', xl_sheet.cell_value(rx, 0))], limit=1)
            # if not brand:
            #     brand_id = self._create_obj('product.brand', xl_sheet.cell_value(rx, 0)).id
            # else:
            #     brand_id = brand.id

            internal_category=self.get_internal_category(str(xl_sheet.cell_value(rx, 4)))
            #pos_category=self._create_obj('pos.category',str(xl_sheet.cell_value(rx, 8))).id


            product_template = self._find_product([('name', '=', name)])
            if product_template:
                # if len(xl_sheet.cell_value(rx, 1)) != 0:
                #     attribute_id = self.env['product.attribute'].search([('name', '=', u"الموسم")], limit=1).id
                #     self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 1))

                if len(str(xl_sheet.cell_value(rx, 0))) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"الموديل")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 0)))

                # if len(xl_sheet.cell_value(rx, 3)) != 0:
                #     attribute_id = self.env['product.attribute'].search([('name', '=', u"السروال")], limit=1).id
                #     self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 3))

                if len(xl_sheet.cell_value(rx, 2)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"اللون")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 2))

                if xl_sheet.cell_value(rx, 3):
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"المقاس")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 3)) )


            if not product_template:
                values = {}

                # if brand_id:
                #     values['product_brand_id'] = brand_id

                if internal_category:
                    values['categ_id'] = internal_category

                # if pos_category:
                #     values['pos_categ_id'] = pos_category

                product_template=self._create_product(name, values=values)
                # if len(xl_sheet.cell_value(rx, 1)) != 0:
                #     attribute_id = self.env['product.attribute'].search([('name', '=', u"الموسم")], limit=1).id
                #     self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 1))

                if len(str(xl_sheet.cell_value(rx, 0))) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"الموديل")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 0)))

                # if len(xl_sheet.cell_value(rx, 3)) != 0:
                #     attribute_id = self.env['product.attribute'].search([('name', '=', u"السروال")], limit=1).id
                #     self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 3))

                if len(xl_sheet.cell_value(rx, 2)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"اللون")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 2))

                if xl_sheet.cell_value(rx, 3):
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"المقاس")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 3)) )

                product_template.generate_barcode()

            logging.info('%s rows remaining' % remaining)

    @api.multi
    def import_product_sultan_zbounet(self):
        book = xlrd.open_workbook(os.path.dirname(__file__) + "/Sultan Products.xlsx")
        xl_sheet = book.sheet_by_index(5)
        remaining = xl_sheet.nrows


        for rx in range(1, xl_sheet.nrows):
            # if row has no product skip
            remaining -= 1
            if not xl_sheet.cell_value(rx, 2):
                continue
            name = xl_sheet.cell_value(rx, 2)

            # brand_obj = self.env["product.brand"]
            # brand = brand_obj.search([('name', '=', xl_sheet.cell_value(rx, 0))], limit=1)
            # if not brand:
            #     brand_id = self._create_obj('product.brand', xl_sheet.cell_value(rx, 0)).id
            # else:
            #     brand_id = brand.id

            internal_category=self.get_internal_category(str(xl_sheet.cell_value(rx, 5)))
            #pos_category=self._create_obj('pos.category',str(xl_sheet.cell_value(rx, 8))).id


            product_template = self._find_product([('name', '=', name)])
            if product_template:
                if len(xl_sheet.cell_value(rx, 0)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"نوعية الشغل")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 0))

                if len(str(xl_sheet.cell_value(rx, 1))) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"نوعية القماش")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 1)))

                # if len(xl_sheet.cell_value(rx, 3)) != 0:
                #     attribute_id = self.env['product.attribute'].search([('name', '=', u"السروال")], limit=1).id
                #     self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 3))

                if len(xl_sheet.cell_value(rx, 3)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"اللون")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 3))

                if xl_sheet.cell_value(rx, 4):
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"المقاس")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 4)) )


            if not product_template:
                values = {}

                # if brand_id:
                #     values['product_brand_id'] = brand_id

                if internal_category:
                    values['categ_id'] = internal_category

                # if pos_category:
                #     values['pos_categ_id'] = pos_category

                product_template=self._create_product(name, values=values)
                if len(xl_sheet.cell_value(rx, 0)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"نوعية الشغل")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 0))

                if len(str(xl_sheet.cell_value(rx, 1))) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"نوعية القماش")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 1)))

                # if len(xl_sheet.cell_value(rx, 3)) != 0:
                #     attribute_id = self.env['product.attribute'].search([('name', '=', u"السروال")], limit=1).id
                #     self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 3))

                if len(xl_sheet.cell_value(rx, 3)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"اللون")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 3))

                if xl_sheet.cell_value(rx, 4):
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"المقاس")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 4)) )

                product_template.generate_barcode()

            logging.info('%s rows remaining' % remaining)

    @api.multi
    def import_product_sultan_badla(self):
        book = xlrd.open_workbook(os.path.dirname(__file__) + "/Sultan Products.xlsx")
        xl_sheet = book.sheet_by_index(8)
        remaining = xl_sheet.nrows


        for rx in range(1, xl_sheet.nrows):
            # if row has no product skip
            remaining -= 1
            if not xl_sheet.cell_value(rx, 2):
                continue
            name = xl_sheet.cell_value(rx, 2)

            # brand_obj = self.env["product.brand"]
            # brand = brand_obj.search([('name', '=', xl_sheet.cell_value(rx, 0))], limit=1)
            # if not brand:
            #     brand_id = self._create_obj('product.brand', xl_sheet.cell_value(rx, 0)).id
            # else:
            #     brand_id = brand.id

            internal_category=self.get_internal_category(str(xl_sheet.cell_value(rx, 6)))
            #pos_category=self._create_obj('pos.category',str(xl_sheet.cell_value(rx, 8))).id


            product_template = self._find_product([('name', '=', name)])
            if product_template:
                if len(xl_sheet.cell_value(rx, 0)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"نوعية الشغل")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 0))

                if len(str(xl_sheet.cell_value(rx, 1))) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"نوعية القماش")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 1)))

                if len(xl_sheet.cell_value(rx, 3)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"السروال")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 3))

                if len(xl_sheet.cell_value(rx, 4)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"اللون")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 4))

                if xl_sheet.cell_value(rx, 5):
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"المقاس")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 5)) )


            if not product_template:
                values = {}

                # if brand_id:
                #     values['product_brand_id'] = brand_id

                if internal_category:
                    values['categ_id'] = internal_category

                # if pos_category:
                #     values['pos_categ_id'] = pos_category

                product_template=self._create_product(name, values=values)
                if len(xl_sheet.cell_value(rx, 0)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"نوعية الشغل")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 0))

                if len(str(xl_sheet.cell_value(rx, 1))) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"نوعية القماش")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 1)))

                if len(xl_sheet.cell_value(rx, 3)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"السروال")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 3))

                if len(xl_sheet.cell_value(rx, 4)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"اللون")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 4))

                if xl_sheet.cell_value(rx, 5):
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"المقاس")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 5)) )

                product_template.generate_barcode()

            logging.info('%s rows remaining' % remaining)


    @api.multi
    def import_product_sultan_farmla(self):
        book = xlrd.open_workbook(os.path.dirname(__file__) + "/Sultan Products.xlsx")
        xl_sheet = book.sheet_by_index(9)
        remaining = xl_sheet.nrows


        for rx in range(1, xl_sheet.nrows):
            # if row has no product skip
            remaining -= 1
            if not xl_sheet.cell_value(rx, 2):
                continue
            name = xl_sheet.cell_value(rx, 2)

            # brand_obj = self.env["product.brand"]
            # brand = brand_obj.search([('name', '=', xl_sheet.cell_value(rx, 0))], limit=1)
            # if not brand:
            #     brand_id = self._create_obj('product.brand', xl_sheet.cell_value(rx, 0)).id
            # else:
            #     brand_id = brand.id

            internal_category=self.get_internal_category(str(xl_sheet.cell_value(rx, 6)))
            #pos_category=self._create_obj('pos.category',str(xl_sheet.cell_value(rx, 8))).id


            product_template = self._find_product([('name', '=', name)])
            if product_template:
                if len(xl_sheet.cell_value(rx, 0)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"نوعية الشغل")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 0))

                if len(str(xl_sheet.cell_value(rx, 1))) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"نوعية القماش")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 1)))

                if len(xl_sheet.cell_value(rx, 3)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"السروال")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 3))

                if len(xl_sheet.cell_value(rx, 4)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"اللون")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 4))

                # if xl_sheet.cell_value(rx, 5):
                #     attribute_id = self.env['product.attribute'].search([('name', '=', u"المقاس")], limit=1).id
                #     self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 5)) )


            if not product_template:
                values = {}

                # if brand_id:
                #     values['product_brand_id'] = brand_id

                if internal_category:
                    values['categ_id'] = internal_category

                # if pos_category:
                #     values['pos_categ_id'] = pos_category

                product_template=self._create_product(name, values=values)
                if len(xl_sheet.cell_value(rx, 0)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"نوعية الشغل")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 0))

                if len(str(xl_sheet.cell_value(rx, 1))) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"نوعية القماش")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 1)))

                if len(xl_sheet.cell_value(rx, 3)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"السروال")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 3))

                if len(xl_sheet.cell_value(rx, 4)) != 0:
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"اللون")], limit=1).id
                    self._add_variant(product_template, attribute_id, xl_sheet.cell_value(rx, 4))

                if xl_sheet.cell_value(rx, 5):
                    attribute_id = self.env['product.attribute'].search([('name', '=', u"المقاس")], limit=1).id
                    self._add_variant(product_template, attribute_id, str(xl_sheet.cell_value(rx, 5)) )

                product_template.generate_barcode()

            logging.info('%s rows remaining' % remaining)


    @api.multi
    def import_all_products(self):
        self.import_product_attribute()
        self.import_product_sultan_badla_arabe()
        self.import_product_sultan_thawb()
        self.import_product_sultan_horka()
        self.import_product_sultan_dakhilia()
        self.import_product_sultan_accessoires()
        self.import_product_sultan_farmla()

