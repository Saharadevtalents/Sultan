# -*- encoding: utf-8 -*-
import logging
from odoo import models, fields, api, _

_logger = logging.getLogger(__name__)

try:
    import barcode
except ImportError:
    _logger.debug("Cannot import 'viivakoodi' python library.")
    barcode = None


class ProductTemplate(models.Model):
    _inherit = "product.template"

    barcode_base = fields.Char()
    sequence_id = fields.Many2one('ir.sequence', ondelete='cascade', string='Sequence')
    pattern_code = fields.Char(readonly=True, default='625.....0000')
    padding = fields.Char(readonly=True, default='5')

    @api.multi
    def _generate_base(self):
        self.sequence_id = self.env.ref('generate_barcode_ean13.seq_barcode_base').id
        if not self.barcode_base:
            self.barcode_base = self.sequence_id.next_by_id()

    @api.multi
    def generate_barcode(self):
        self._generate_base()
        if self.pattern_code and self.padding and self.barcode_base:
            str_barcode_base = str(self.barcode_base).rjust(int(self.padding), '0')
            pattern_code = self.pattern_code.replace('.' * int(self.padding), str_barcode_base)
            barcode_class = barcode.get_barcode_class('ean13')
            self.barcode = barcode_class(pattern_code)

    @api.multi
    def generate_barcode_server(self):
        for product in self.search([]):
            product.generate_barcode()
