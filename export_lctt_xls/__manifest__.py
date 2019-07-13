{
    'name': 'Export lctt Info in Excel',
    'version': '0.2',
    'category': 'LcTt',
    'license': "AGPL-3",
    'author': 'Itech reosurces',
    'company': 'ItechResources',
    'depends': [
                'base',
                'purchase',
                'report_xlsx',
                'account',
                ],
    'data': [
            'views/wizard_view.xml',

            ],
    'installable': True,
    'auto_install': False,
}
