# -*- coding: utf-8 -*-

from odoo import api, fields, models


class ProductProduct(models.Model):
    _inherit = 'product.product'

    def _get_size(self):
        for attribute_line in self.attribute_line_ids:
            if attribute_line.attribute_id.name == u"المقاس":
                return attribute_line.value_ids[0].name

    def _get_color(self):
        for attribute_line in self.attribute_line_ids:
            if attribute_line.attribute_id.name == u"اللون":
                return attribute_line.value_ids[0].name