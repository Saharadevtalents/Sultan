# -*- coding: utf-8 -*-

{
    'name': 'Product import',
    'version': '1.1',
    'category': 'Product',
    'author': '',
    'sequence': 1,
    'depends': [
        'product',
        'product_brand',
    ],
    "data": [
        'views/product_view.xml',
        'views/product_importer_view.xml',
        'views/product_update_view.xml'
    ],
    'qweb': [
    ],
    'installable': True,
    'auto_install': False,
}
