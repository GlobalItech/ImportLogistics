{
    'name': 'Purchase/ Import with Landed Cost Module (Average Costing)',
    'version': '0.2',
    'category': 'accouting',
    'license': "AGPL-3",
    'summary': " ",
    'author': 'Itech Resources',
    'company': 'ItechResources',
    'depends': [
                'account',
                'purchase',
#                 'obro_workflows',
                ],
    'data': [
#             'security/ir.model.access.csv',


            'data/lc_sequence.xml',
            'views/account_invoice.xml',
            'views/lc_view.xml',
            
            
            ],
    'installable': True,
    'auto_install': False,
    'price':'170.0',
    'currency': 'EUR',
    
    
}
