from odoo.addons.report_xlsx.report.report_xlsx import ReportXlsx
from odoo import fields, models,api
from datetime import datetime, time, date

class lcttReportXls(ReportXlsx):

    @api.multi
    def get_lines(self,date_from,date_to):
        
        lines = []
        purchase_obj_ids = self.env['purchase.order'].search([('date_order', '>=',date_from),
                                                                ('date_order', '<=',date_to),
                                                                ('s_for', '=','import')])
       
        if purchase_obj_ids:
            

            
            for purchase_obj in purchase_obj_ids:
                sum_unit_pricefc =0.000
                for po_line in purchase_obj.order_line:
                    sum_unit_pricefc += po_line.unit_pricefc
                lc_amount=0.0
                for lc_ob in purchase_obj.lc_ids:
                    lc_amount += lc_ob.amount
                    
                    
                
#                 unit_price_fc =[l.unit_pricefc for l in purchase_obj.order_line]
#                 sub_total_lp = [l.sub_total_lp for l in purchase_obj.order_line],
                Insurance = purchase_obj.lc_ids.filtered(lambda x: x.name.name == 'Insurance').amount or 0.0,
                Bank_Charges = purchase_obj.lc_ids.filtered(lambda x: x.name.name == 'Bank Charges').amount or 0.0,
                Govt_Tax_Duties =purchase_obj.lc_ids.filtered(lambda x: x.name.name == 'Govt Tax/Duties').amount or 0.0,
                Demerage = purchase_obj.lc_ids.filtered(lambda x: x.name.name == 'Demerage').amount or 0.0,
                Detensions =purchase_obj.lc_ids.filtered(lambda x: x.name.name == 'Detensions').amount or 0.0,
                Clearing= purchase_obj.lc_ids.filtered(lambda x: x.name.name == 'Clearing').amount or 0.0,
                Freight =purchase_obj.lc_ids.filtered(lambda x: x.name.name == 'Freight').amount or 0.0,
#                 'Cost':0.0
             
                if sum_unit_pricefc !=0:
                    vals = ({
                            'po_name':purchase_obj.name,
                            'ref-no':purchase_obj.lc_ref,
                            'name': purchase_obj.partner_id.name,
                            'bank_name':purchase_obj.bank_name.name,
                            'lc_ref_no':purchase_obj.lc_ref_no,
                            'condition':purchase_obj.condition.name,
                            'date':purchase_obj.date_order,
        #                     'date':datetime.strftime(purchase_obj.date_order, "%Y-%m-%d"),
                            'particular':[l.product_id.name for l in purchase_obj.order_line],
                            'qty':[l.product_qty for l in purchase_obj.order_line],
                            'rate':[l.unit_pricefc for l in purchase_obj.order_line],
                            'value':[l.sub_total_fc for l in purchase_obj.order_line],
                            'fx_rate':purchase_obj.fx_rate,
        #                     p.amount = subtotal
                            'sub_total_lp':[l.sub_total_lp for l in purchase_obj.order_line],
                            
                            'Insurance':(Insurance[0]/sum_unit_pricefc) or 0.0,
                            'Bank Charges':(Bank_Charges[0]/sum_unit_pricefc) or 0.0,
                            'Govt Tax/Duties':(Govt_Tax_Duties[0]/sum_unit_pricefc) or 0.0,
                            'Demerage': (Demerage[0]/sum_unit_pricefc) or 0.0,
                            'Detensions':(Detensions[0]/sum_unit_pricefc) or 0.0,
                            'Clearing':(Clearing[0]/sum_unit_pricefc) or 0.0,
                            'Freight':(Freight[0]/sum_unit_pricefc)  or 0.0,
                            'Cost':0.0
                       })
                    lines.append(vals)
                            
                else:
                    vals = ({
                        'po_name':purchase_obj.name,
                        'ref-no':purchase_obj.lc_ref,
                        'name': purchase_obj.partner_id.name,
                        'bank_name':purchase_obj.bank_name.name,
                        'lc_ref_no':purchase_obj.lc_ref_no,
                        'condition':purchase_obj.condition.name,
                        'date':purchase_obj.date_order,
    #                     'date':datetime.strftime(purchase_obj.date_order, "%Y-%m-%d"),
                        'particular':[l.product_id.name for l in purchase_obj.order_line],
                        'qty':[l.product_qty for l in purchase_obj.order_line],
                        'rate':[l.unit_pricefc for l in purchase_obj.order_line],
                        'value':[l.sub_total_fc for l in purchase_obj.order_line],
                        'fx_rate':purchase_obj.fx_rate,
    #                     p.amount = subtotal
                        'sub_total_lp':[l.sub_total_lp for l in purchase_obj.order_line],
                        
                        'Insurance':purchase_obj.lc_ids.filtered(lambda x: x.name.name == 'Insurance').amount or 0.0,
                        'Bank Charges':purchase_obj.lc_ids.filtered(lambda x: x.name.name == 'Bank Charges').amount or 0.0,
                        'Govt Tax/Duties':purchase_obj.lc_ids.filtered(lambda x: x.name.name == 'Govt Tax/Duties').amount or 0.0,
                        'Demerage': purchase_obj.lc_ids.filtered(lambda x: x.name.name == 'Demerage').amount or 0.0,
                        'Detensions':purchase_obj.lc_ids.filtered(lambda x: x.name.name == 'Detensions').amount or 0.0,
                        'Clearing':purchase_obj.lc_ids.filtered(lambda x: x.name.name == 'Clearing').amount or 0.0,
                        'Freight':purchase_obj.lc_ids.filtered(lambda x: x.name.name == 'Freight').amount or 0.0,
                        'Cost':0.0
                        })
                    lines.append(vals)
 
 
        return lines



    def generate_xlsx_report(self, workbook, data, lines):
        sheet = workbook.add_worksheet()
        format1 = workbook.add_format({'font_size': 14, 'bottom': True, 'right': True, 'left': True, 'top': True, 'align': 'vcenter', 'bold': True})
        format11 = workbook.add_format({'font_size': 12, 'align': 'center', 'right': True, 'left': True, 'bottom': True, 'top': True, 'bold': True})
        format21 = workbook.add_format({'font_size': 10, 'align': 'center', 'right': True, 'left': True,'bottom': True, 'top': True, 'bold': True})
        format21.set_num_format('#,##0.00')
        format3 = workbook.add_format({'bottom': True, 'top': True, 'font_size': 12})
        font_size_8 = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 8})
        red_mark = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 8,
                                        'bg_color': 'red'})
        justify = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 12})
#         style = workbook.add_format('align: wrap yes; borders: top thin, bottom thin, left thin, right thin;')
#         style.num_format_str = '#,##0.00'
        format3.set_align('center')
        font_size_8.set_align('center')
        justify.set_align('justify')
        format1.set_align('center')
        red_mark.set_align('center')
        
        date_from = datetime.strptime(data['form']['date_from'], '%Y-%m-%d').strftime('%d/%m/%y')
        date_to = datetime.strptime(data['form']['date_to'], '%Y-%m-%d').strftime('%d/%m/%y')
        sheet.merge_range(1, 0, 3, 14, 'LCTT Record', format1)
        sheet.merge_range(4, 0, 5, 14, 'Period from :' + (date_from) +  ' To ' + (date_to), format1)
#         

        sheet.write(10, 0, 'PO No', format21)
        sheet.write(10, 1, 'Ref No', format21)
        sheet.write(10, 2, 'LC No', format21)
        sheet.write(10, 3, 'Supplier', format21)
        sheet.write(10, 4, 'Bank Names', format21)
        sheet.write(10, 5, 'Condition', format21)
        sheet.write(10, 6, 'Date', format21)
        sheet.write(10, 7, 'Particular', format21)
        sheet.write(10, 8, 'QTY', format21)
        sheet.write(10, 9 , 'Rate ($)', format21)
        sheet.write(10, 10, 'Value ($)', format21)
        sheet.write(10, 11, 'Ex. Rate', format21)
        sheet.write(10, 12, 'Principle Amount', format21)
        sheet.write(10, 13, 'Insurance', format21)
        sheet.write(10, 14, 'Bank Charges', format21)
        sheet.write(10, 15, 'Govt Tax/Duties', format21)
        sheet.write(10, 16, 'Demerage', format21)
        sheet.write(10, 17, 'Detensions', format21)
        sheet.write(10, 18, 'Clearing', format21)
        sheet.write(10, 19, 'Freight', format21)
        sheet.write(10, 20, 'Net', format21)
        sheet.write(10, 21, 'Cost', format21)
        # report statrt
        product_row = 13
        get_lines = self.get_lines(data['form']['date_from'],data['form']['date_to'])
       
        c=0
        sr_no =0
        for line in  get_lines:
#             date =datetime.strptime(line['date'], '%Y-%m-%d').strftime('%d/%m/%y')
   
            for ln in line:
                
                if c<len(line['particular']):
                    sheet.write(product_row, 0, line['po_name'], format21)
                    sheet.write(product_row, 1, line['ref-no'], format21)
                    sheet.write(product_row, 2, line['lc_ref_no'], format21)
                    sheet.write(product_row, 3, line['name'], format21)
                    sheet.write(product_row, 4, line['bank_name'], format21)
                    sheet.write(product_row, 5, line['condition'], format21)
                    sheet.write(product_row, 6, line['date'], format21)
                    sheet.write(product_row, 7, line['particular'][c], format21)
                    sheet.write(product_row, 8, line['qty'][c], format21)
                    sheet.write(product_row, 9, line['rate'][c], format21)
                    sheet.write(product_row, 10, line['value'][c], format21)
                    sheet.write(product_row, 11, line['fx_rate'], format21)
                    sheet.write(product_row, 12, line['sub_total_lp'][c], format21)
                    sheet.write(product_row, 13,  line['Insurance']* line['rate'][c], format21)
                    sheet.write(product_row, 14,  line['Bank Charges']* line['rate'][c], format21)
                    sheet.write(product_row, 15,  line['Govt Tax/Duties']* line['rate'][c], format21)
                    sheet.write(product_row, 16,  line['Demerage']* line['rate'][c], format21)
                    sheet.write(product_row, 17,  line['Detensions']* line['rate'][c], format21)
                    sheet.write(product_row, 18,  line['Clearing']* line['rate'][c], format21)
                    sheet.write(product_row, 19,  line['Freight']* line['rate'][c], format21)
                    net =0.0
#                     net = line['sub_total_lp'][c] + line['Insurance']+line['Bank Charges']+line['Govt Tax/Duties']+line['Demerage']+ line['Detensions'] +line['Clearing']+ line['Freight'] or 0.0
                    net = line['sub_total_lp'][c] + line['Insurance']* line['rate'][c] +line['Bank Charges']* line['rate'][c] + line['Govt Tax/Duties']* line['rate'][c]+line['Demerage']* line['rate'][c]+ line['Detensions'] * line['rate'][c] +line['Clearing'] * line['rate'][c]+ line['Freight']* line['rate'][c] or 0.0

                    sheet.write(product_row, 20, net, format21)
                    if line['qty'][c] !=0:
                        sheet.write(product_row, 21, net/line['qty'][c], format21)

                    product_row +=2
                c +=1
                
                
            
            c=0
            sr_no +=1   
            
            
        

            
lcttReportXls('report.export_lctt_xls.lctt_report_xls.xlsx','purchase.order')
