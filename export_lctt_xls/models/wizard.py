from openerp import models, fields, api


class StockReport(models.TransientModel):
    _name = "wizard.lctt.history"
    _description = "Current Stock History"

    date_to= fields.Date("Date To")
    date_from= fields.Date("Date From")
#     report_type = fields.Selection([('product_wise','Product Wise Report'),('grand_summary','Grand Summary'),('sale_partywise','Party Wise')],string='Relative')
#     category = fields.Many2many('product.category', 'categ_wiz_rel', 'categ', 'wiz', string='Warehouse')

    @api.multi
    def export_xls(self):
        context = self._context
        datas = {'ids': context.get('active_ids', [])}
        datas['model'] = 'purchase.order'
        datas['form'] = self.read()[0]
        for field in datas['form'].keys():
            if isinstance(datas['form'][field], tuple):
                datas['form'][field] = datas['form'][field][0]
        if context.get('xls_export'):
            return {'type': 'ir.actions.report.xml',
                    'report_name': 'export_lctt_xls.lctt_report_xls.xlsx',
                    'datas': datas,
                    'name': 'lctt'
                    }
          