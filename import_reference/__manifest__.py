# -*- coding: utf-8 -*-

{
    'name': 'Import reference',
    'version' : '1.0',
    'author' : 'DHOUIBI Dorsaf',
    'website' : ' ',
    'category' : 'Tools',
    'depends' : ['base','auto_parts'],
    'description': '''
          Import preference from Excel file
     ''',
    'data' : [
            'security/ir.model.access.csv',
            'views/import_reference_view.xml',
        ],
    'installable' : True,
}