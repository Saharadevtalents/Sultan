# -*- coding: utf-8 -*-
{
    'name': 'Generate Barcode EAN 13',
    'version': '1.1',
    'sequence': 2,
    'author': 'DevTalents',
    'summary': 'Generate Barcode EAN 13',
    'description': """
    Generate Barcode EAN 13
        """,
    'depends': ['sale_management'],
    'data': [
        'data/ir_sequence_data.xml',
        'views/product_template.xml',
    ],
    'qweb': [
    ],
    'installable': True,
    'auto_install': False,
    'application': True,
}
