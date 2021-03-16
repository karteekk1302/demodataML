from odoo import models, api, _
import base64
import xlrd
from datetime import datetime, timedelta
from odoo.exceptions import AccessError
from odoo.exceptions import AccessError
from odoo.exceptions import AccessError
from odoo.exceptions import AccessError


class ticl_shipment_log(models.Model):
    _name = 'ticl.tender.import'
    

    def getSipmentType(self,equipment_date):
        eqp_dt_time = datetime.strptime(equipment_date, '%Y-%m-%d %H:%M:%S')
        diff_sev = eqp_dt_time.date() - timedelta(days=7)
        time_between_insertion = diff_sev - datetime.now().date()
        if  time_between_insertion.days >= 5: shipType = 'regular'
        else: shipType = 'expedited'
        return shipType
    
    @api.model
    def shipment_tender_import_ext(self, vals):
        message = 'Tender Imported Successfully!'
        status = 's'
        product_list = []
        try:
            xl = vals.get('file').split(',')
            xlsx_file = xl[1].encode()
            xls_file = base64.decodestring(xlsx_file)
            wb = xlrd.open_workbook(file_contents=xls_file)
            
            for sheet in wb.sheets():
                shipTo = sheet.col_values(11)
                shipTo.pop(0)
                locations = self.env['stock.location'].sudo().search([('name','in',shipTo)]).ids
                for location in locations:
                    lst = []
                    vals = {}
                    warehouse = {}
                    for row in range(sheet.nrows):
                        if row == 0:
                            continue
                        dt = {}
                        location_name = sheet.cell(row,11).value
                        loc = self.env['stock.location'].sudo().search([('name','=',location_name)],limit=1)
                        if not loc:
                            exception = AccessError("No value found for row - %s and column - %s" % (row, 11))
                            raise exception
#                             self._cr.rollback()
#                             message = str(location_name) + ' not found contact with your administrator.'
#                             return {'message':message}
                        if location == loc.id:
                            vals.update({'state':'draft','sending_location_id':loc.id})
                            product_name = str(sheet.cell(row,12).value)
                            if not product_name:
                                exception = AccessError("Verify your models before import.")
                                raise exception
                            if "." in product_name:
                                prod_lst = product_name.split('.')
                                product_name = prod_lst[0]
                            product = self.env['product.product'].sudo().search([('name','=',product_name)],limit=1)
                            if not product:
                                product_list.append(product_name)
                            serial_number = str(sheet.cell(row,4).value)
                            if "." in serial_number:
                                ser_number_lst = serial_number.split('.')
                                serial_number = ser_number_lst[0]
                            for col in range(sheet.ncols):
                                if col in (0,2,3,10,13,15,16):
                                    continue
                                elif col == 4:
                                    dt.update({'serial_number':serial_number})
                                elif col == 5:
                                    dt.update({'funding_doc_type':sheet.cell(row,col).value})
                                elif col == 6:
                                    dt.update({'funding_doc_number':sheet.cell(row,col).value})
                                elif col == 7:
                                    dt.update({'ticl_project_id':sheet.cell(row,col).value})
                                elif col == 8:
                                    dt.update({'tid':sheet.cell(row,col).value})
                                elif col == 9:
                                    dt.update({'common_name':sheet.cell(row,col).value})
                                elif col == 12:
                                    dt.update({
                                        'product_id':product.id,
                                        'tel_type':product.categ_id.id,
                                        'manufacturer_id':product.manufacturer_id.id,
                                        'count_number':1
                                        })
                                    
                                elif col == 1:
                                    eqpDate = datetime(*xlrd.xldate_as_tuple(sheet.cell(row,col).value, wb.datemode))
                                    eqpDay, eqpMonth, eqpYr = eqpDate.day, eqpDate.month, eqpDate.year
                                    eqp_grp = str(eqpDay)+'_'+str(eqpMonth)+'_'+str(eqpYr)
                                    dt.update({'activity_date':eqpDate.strftime("%Y-%m-%d %H:%M:%S"),'eqp_grp':eqp_grp})
                                    eqpDtTime = eqpDate.strftime("%Y-%m-%d %H:%M:%S")
                                    shipment_type = 'regular'
                                    vals.update({
                                        'activity_date':eqpDtTime,
                                        'shipment_types':shipment_type
                                        })
                                elif col == 16:#equipment_date
                                    eqpDate_shp = datetime(*xlrd.xldate_as_tuple(sheet.cell(row,col).value, wb.datemode))
                                    eqpDtTime_shp = eqpDate_shp.strftime("%Y-%m-%d %H:%M:%S")
                                    
                                    vals.update({
                                        'equipment_date':eqpDtTime_shp
                                        })  
                            
                            move = self.env['stock.move'].sudo().search([
                                ('product_id','=',product.id),
                                ('status','=','inventory'),
                                ('serial_number','=',serial_number)
                                ],limit=1)
                            if move:
                                dt.update({'tel_available':'Y'})
                                if move.location_dest_id.id in warehouse:
                                    ware_lst = warehouse.get(move.location_dest_id.id)
                                    ware_lst.append(dt)
                                else:
                                    warehouse.update({move.location_dest_id.id:[dt]})
                            else:
                                dt.update({'tel_available':'N'})
                                lst.append((0,0,dt))
                            
                    if product_list:
                        msg = ','.join(product_list)
                        message = msg + ' not in products! check your products.'
                        exception = AccessError(message)
                        raise exception
                    if warehouse:
                        for ware in warehouse.keys():
                            
                            eqp_dict = {}
                            for data in warehouse.get(ware):
#                                 warehouse_list.append((0,0,data))
                                if data.get('eqp_grp') in eqp_dict:
                                    eqp_lst = eqp_dict.get(data.get('eqp_grp'))
                                    eqp_lst.append(data)
                                else:
                                    eqp_dict.update({data.get('eqp_grp'):[data]})
                            for eqp_key in eqp_dict.keys():
                                warehouse_list = []
                                activity_date = ''
                                for tender_data in eqp_dict.get(eqp_key):
                                    warehouse_list.append((0,0,tender_data))
                                    activity_date = tender_data.get('activity_date')
                                    shipment_type = 'regular'
                                    vals.update({'shipment_types':shipment_type})
                                vals.update({
                                    'ticl_ship_lines':warehouse_list,
                                    #'warehouse_id':int(ware),
                                    'receiving_location_id':int(ware),
                                    'activity_date':activity_date
                                })
                                avl_shp = self.env['ticl.shipment.log.ext'].create(vals)
                                #avl_shp.picked_shipment_log_ext()
                                if lst:
                                    vals.update({'ticl_ship_lines':lst})
                                    ship_log = self.env['ticl.shipment.log.ext'].create(vals)
                                    pending = ship_log.ticl_ship_lines.filtered(lambda p: p.tel_available == 'N')
                                    if not pending:
                                        pass
                                    else:
                                        ship_log.picked_shipment_log_ext()
                                    lst = []
                                else:
                                    avl_shp.picked_shipment_log_ext()
                                    if avl_shp.state == 'approved':
                                        for ticl_ship_line in avl_shp.ticl_ship_lines:
                                            move_inv = self.env['stock.move'].search(
                                                [['serial_number', '=', ticl_ship_line.serial_number],
                                                 ['status', '=', 'inventory']],limit=1)
                                            if move_inv:
                                                move_inv.status = 'assigned'
                                                ticl_ship_line.move_id = move_inv.id


                    else:
                        vals.update({'ticl_ship_lines':lst})
                        ship_log = self.env['ticl.shipment.log.ext'].create(vals)
                        pending = ship_log.ticl_ship_lines.filtered(lambda p: p.tel_available == 'N')
                        if not pending:
                            pass
                        else:
                            ship_log.picked_shipment_log_ext()
        except Exception as e:
            status = 'n'
            message = str(e)
        return {'message':message,'status':status}
