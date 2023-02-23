from openpyxl import load_workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
import pyexcel
import os
from datetime import datetime

months = {'01':'январь', '02':'февраль', '03':'март', '04':'апрель',\
         '05':'май', '06':'июнь', '07':'июль', '08':'август',\
         '09':'сеньтябрь', '10':'октябрь', '11':'ноябрь', '12':'декабрь'}
def get_month(m='01-01-2023'):
    return months[m[3:5]]


def get_head_data(my_dir, factory_num, inp_type):
    header_ws = load_workbook(my_dir + '\Templates\HeadDat.xlsx', data_only=False).active
    head_data = {'factory_num': '', 'complex_num': '', 'consumer': '', 'order': '', 'adress': '', 'cold_temp': '5,0', 'save_folder': ''}
    for i in range(2, 426):
        if  str(header_ws['A' + str(i)].value) in factory_num + '_' + inp_type:
            head_data['factory_num'] = factory_num
            head_data['complex_num'] = header_ws['B' + str(i)].value
            head_data['consumer'] = header_ws['C' + str(i)].value
            head_data['order'] = header_ws['D' + str(i)].value
            head_data['adress'] = header_ws['E' + str(i)].value
            head_data['cold_temp'] = header_ws['G' + str(i)].value
            head_data['save_folder'] = header_ws['J' + str(i)].value
        elif '-' in factory_num and str(header_ws['A' + str(i)].value.split('_')[0]) in factory_num + '_' + inp_type:
            head_data['factory_num'] = factory_num
            head_data['complex_num'] = header_ws['B' + str(i)].value
            head_data['consumer'] = header_ws['C' + str(i)].value
            head_data['order'] = header_ws['D' + str(i)].value
            head_data['adress'] = header_ws['E' + str(i)].value
            head_data['cold_temp'] = header_ws['G' + str(i)].value
            head_data['save_folder'] = header_ws['J' + str(i)].value
    
    return head_data


class MKTSParser:
    def __init__(self, data_list, curr_dir, save_dir):
        self.my_parsing_files = []
        self.my_dir = curr_dir
        self.save_dir = save_dir
        for file in data_list['МКТС']:
            if 'xlsx' not in file:
                pyexcel.save_book_as(file_name=file,
                                dest_file_name=file + 'x')
                file += 'x'
            self.my_parsing_files.append([file, load_workbook(filename=file,  read_only=True).active])


    def data_index(self, row):
        nums_of_data = {'Дат': -1, 't1,°С': -1, 't2,°С': -1,'V1': -1,'M1': -1,'V2': -1,'M2': -1, 'Q': -1, 'Tраб': -1, 'Tотк': -1, 'Отка': -1}
        for key in nums_of_data.keys():
            ind = 0
            for cell in row:
                if key in cell.value:
                    nums_of_data[key] = ind
                    break
                ind += 1

        return nums_of_data


    def __call__(self, date_from = '01-01-2023', date_to = '18-01-2023'):
        report = '\tМКТС:\n' # Window print
        for file in self.my_parsing_files:
            template = load_workbook(self.my_dir + '\Templates\VEC_Template.xlsx',  read_only=False, data_only=False)  # Template xlsx file  
            file_name = file[0].split('/')[len(file[0].split('/')) - 1].split('.xlsx')[0]
            if file[0].split('/')[len(file[0].split('/')) - 1]:
                template.title = file[0].split('/')[len(file[0].split('/')) - 1]
            ws = template.active

            order_type = '1'
            if 'ГВС' in file[0]:
                order_type = '2'

            head_data = {}
            data_index = {}
            row_index = 1
            out_index = 18

            dm_sum = 0
            m1_sum = 0
            m2_sum = 0 
            dv_sum = 0
            v1_sum = 0
            v2_sum = 0 
            t1_avg = 0
            t2_avg = 0
            q_sum = 0
            vnr = 0
            vos = 0        
            sum_err = ''    
            for row in file[1].iter_rows():
                num = lambda t: round(float(row[data_index[t]].value), 2)
                st_row = lambda n: str(n).replace('.', ',')
                if row_index == 1: 
                    head_data = get_head_data(self.my_dir, row[0].value.split('№')[1].split(',')[0], order_type)
                    row_index += 1
                    continue
                if row_index == 2:
                    data_index = self.data_index(row)
                    row_index += 1
                    continue
                if 'Итого' in row[0].value:
                    break
    
                tmp_date = row[0].value.split(',')
                curr_date = datetime.strptime(tmp_date[0] + '-' + tmp_date[1] + '-' + '20' + tmp_date[2], "%d-%m-%Y").date()
                if (curr_date >= datetime.strptime(date_from, "%d-%m-%Y").date() and\
                    curr_date <= datetime.strptime(date_to, "%d-%m-%Y").date()):
                    ws.insert_rows(out_index)
                    for i in range(1, 14):
                        thin = Side(border_style="thin", color="000000")
                        ws.cell(out_index, i).border = Border(top=thin, left=thin, right=thin, bottom=thin)
                        ws.cell(out_index, i).alignment = Alignment(horizontal="center", vertical="center")
                    #[print(row[i].value, ' ') for i in range(0, 15)]
                    ws['A' + str(out_index)] = str(curr_date.strftime("%d-%m-%Y"))
                    if data_index['t1,°С'] == -1 or '-' in str(row[data_index['t1,°С']].value):
                        ws['B' + str(out_index)] = ' - '
                    else:
                        ws['B' + str(out_index)] = st_row(num('t1,°С'))
                        t1_avg += num('t1,°С')

                    if data_index['t2,°С'] == -1 or '-' in str(row[data_index['t2,°С']].value):
                        ws['C' + str(out_index)] = ' - '
                    else:
                        ws['C' + str(out_index)] = st_row(num('t2,°С'))
                        t2_avg += num('t2,°С')

                    if data_index['V1'] == -1 or '-' in str(row[data_index['V1']].value):
                        ws['D' + str(out_index)] = ' - '
                    else:
                        ws['D' + str(out_index)] = st_row(num('V1'))
                        v1_sum += num('V1')
                        ws['D' + str(out_index + 1)] = st_row(round(v1_sum))
                    if data_index['M1'] == -1 or '-' in str(row[data_index['M1']].value):
                        ws['E' + str(out_index)] = ' - '
                    else:
                        ws['E' + str(out_index)] = st_row(num('M1'))
                        m1_sum += num('M1')
                        ws['E' + str(out_index + 1)] = st_row(round(m1_sum, 2))
                    if data_index['V2'] == -1 or '-' in str(row[data_index['V2']].value):
                        ws['F' + str(out_index)] = ' - '
                        ws['H' + str(out_index)] = ws['D' + str(out_index)].value
                    else:
                        ws['F' + str(out_index)] = st_row(num('V2'))
                        v2_sum += num('V2')
                        dv_sum += round(abs(num('V2') - num('V1')), 2)
                        ws['F' + str(out_index + 1)] = st_row(round(v2_sum, 2))
                        ws['H' + str(out_index)] = st_row(round(abs(num('V2') - num('V1')), 2))

                    if data_index['M2'] == -1 or '-' in str(row[data_index['M2']].value):
                        ws['G' + str(out_index)] = ' - '
                        ws['I' + str(out_index)] = ws['E' + str(out_index)].value
                    else:
                        ws['G' + str(out_index)] = st_row(num('M2'))
                        m2_sum += num('M2')
                        dv_sum += round(abs(num('M2') - num('M1')), 2)
                        ws['G' + str(out_index + 1)] = st_row(round(m2_sum, 2))
                        ws['I' + str(out_index)] = st_row(round(abs(num('M2') - num('M1')), 2))

                    if data_index['Q'] == -1  or '-' in str(row[data_index['Q']].value):
                        ws['J' + str(out_index)] = ' - '
                    else:
                        ws['J' + str(out_index)] = st_row(num('Q'))
                        q_sum += num('Q')
                        ws['J' + str(out_index + 1)] = st_row(round(q_sum, 2))
                    
                    if data_index['Tраб'] == -1 or '-' in str(row[data_index['Tраб']].value):
                        ws['K' + str(out_index)] = ' - '
                    else:
                        ws['k' + str(out_index)] = st_row(num('Tраб'))
                        vnr += num('Tраб')
                        ws['K' + str(out_index + 1)] = st_row(round(vnr, 2))

                    if data_index['Tотк'] == -1 or '-' in str(row[data_index['Tотк']].value):
                        ws['L' + str(out_index)] = ' - '
                    else:
                        ws['L' + str(out_index)] = st_row(num('Tотк'))
                        vos += num('Tотк')
                        ws['L' + str(out_index + 1)] = st_row(round(vos, 2))

                    if data_index['Отка'] == -1 or type(row[data_index['Отка']].value) != str:
                        ws['M' + str(out_index)] = ' '
                    else:
                        ws['M' + str(out_index)] = row[data_index['Отка']].value
                        if sum_err == '': sum_err += row[data_index['Отка']].value
                        elif row[data_index['Отка']].value not in sum_err:
                            sum_err += ', ' + sum_err + row[data_index['Отка']].value

                    out_index += 1
                    row_index += 1

            if t1_avg != 0:
                ws['B' + str(out_index + 1)] = str(round(t1_avg / (row_index-4), 2)).replace('.', ',')
            if t2_avg != 0:
                ws['C' + str(out_index + 1)] = str(round(t2_avg / (row_index-4), 2)).replace('.', ',')
            
            # A resoult table 
            q_col = 'D'
            vnr_col = 'E'
            vos_col = 'F'
            
            sec_row = out_index + 4
            row_index += 4
            final_m1 = '-'; final_m2 ='-'; final_v1 = '-'; final_v2 = '-'; final_q = '-'; final_vnr = '-'; final_vos = '-'
            if data_index['M1'] != -1:
                if '-' not in str(file[1].cell(row=row_index, column=data_index['M1']).value):
                    final_m1 = float(str(file[1].cell(row=row_index, column=data_index['M1']).value).replace(',', '.').replace(' ', ''))
            if data_index['M2'] != -1:
                if '-' not in str(file[1].cell(row=row_index, column=data_index['M2']).value):
                    final_m2 = float(str(file[1].cell(row=row_index, column=data_index['M2']).value).replace(',', '.').replace(' ', ''))
            if data_index['V1'] != -1:
                if '-' not in str(file[1].cell(row=row_index, column=data_index['V1']).value):
                    final_v1 = float(str(file[1].cell(row=row_index, column=data_index['V1']).value).replace(',', '.').replace(' ', ''))
            else:
                final_v1 = '-'
            if data_index['V2'] != -1: 
                if '-' not in str(file[1].cell(row=row_index, column=data_index['V2']).value):
                    final_v2 = float(str(file[1].cell(row=row_index, column=data_index['V2']).value).replace(',', '.').replace(' ', ''))

            if data_index['Q'] != -1: 
                if '-' not in str(file[1].cell(row=row_index, column=data_index['Q']).value):
                    final_q = float(str(file[1].cell(row=row_index, column=data_index['Q']).value).replace(',', '.').replace(' ', ''))

            if data_index['Tраб'] != -1:
                if '-' not in str(file[1].cell(row=row_index, column=data_index['Tраб']).value):
                    final_vnr = float(str(file[1].cell(row=row_index, column=data_index['Tраб']).value).replace(',', '.').replace(' ', ''))
            if data_index['Tотк'] != -1: 
                if '-' not in str(file[1].cell(row=row_index, column=data_index['Tотк']).value):  
                    final_vos = float(str(file[1].cell(row=row_index, column=data_index['Tотк']).value).replace(',', '.').replace(' ', ''))

            ws['A' + str(sec_row)] = date_from
            ws['A' + str(sec_row + 1)] = date_to
            if final_m1 != '-':
                ws['B' + str(sec_row)] = final_m1
                ws['B' + str(sec_row + 1)] = str(round(m1_sum + final_m1, 2)).replace('.', ',')
            else:
                ws['B' + str(sec_row)] = 0
                ws['B' + str(sec_row + 1)] = str(round(m1_sum, 2)).replace('.', ',')
            if final_m2 != '-':
                ws['C' + str(sec_row)] = 0
                ws['C' + str(sec_row + 1)] = str(round(m2_sum, 2)).replace('.', ',')

            if str(order_type) == '2':
                ws['D' + str(sec_row - 1)] = 'V1, м3'
                if final_v1 != '-':
                    ws['D' + str(sec_row)] = final_v1
                    ws['D' + str(sec_row + 1)] = str(round(v1_sum + final_v1, 2)).replace('.', ',')
                else:
                    ws['D' + str(sec_row)] = 0
                    ws['D' + str(sec_row + 1)] = str(round(v1_sum, 2)).replace('.', ',')
                ws['E' + str(sec_row - 1)] = 'V2, м3'
                if final_v2 != '-':
                    ws['E' + str(sec_row)] = final_v2
                    ws['E' + str(sec_row + 1)] = str(round(v2_sum + final_v2, 2)).replace('.', ',')
                else:
                    ws['E' + str(sec_row)] = 0
                    ws['E' + str(sec_row + 1)] = str(round(v2_sum, 2)).replace('.', ',')
                ws['F' + str(sec_row - 1)] = 'Q, Гкал'
                ws['G' + str(sec_row - 1)] = 'ВНР, час'
                ws['H' + str(sec_row - 1)] = 'ВОС, час'
                q_col = 'F'
                vnr_col = 'G'
                vos_col = 'H'
                for r in range(sec_row - 1, sec_row + 2):
                    thin = Side(border_style="thin", color="000000")
                    ws.cell(r, 7).border = Border(top=thin, left=thin, right=thin, bottom=thin)
                    ws.cell(r, 7).alignment = Alignment(horizontal="center", vertical="center")
                    ws.cell(r, 8).border = Border(top=thin, left=thin, right=thin, bottom=thin)
                    ws.cell(r, 8).alignment = Alignment(horizontal="center", vertical="center")
            
            if final_q != '-':
                ws[q_col + str(sec_row)] = final_q
                ws[q_col + str(sec_row + 1)] = str(round(q_sum + final_q, 2)).replace('.', ',')
            else:
                ws[q_col + str(sec_row)] = 0
                ws[q_col + str(sec_row + 1)] = str(round(q_sum, 2)).replace('.', ',')
            t = final_vos
            final_vos = final_vnr
            final_vnr = t
            if final_vnr != '-':
                ws[vnr_col + str(sec_row)] = final_vnr
                ws[vnr_col + str(sec_row + 1)] = str(round(vnr + final_vnr, 2)).replace('.', ',')
            else:
                ws[vnr_col + str(sec_row)] = 0
                ws[vnr_col + str(sec_row + 1)] = str(round(vnr, 2)).replace('.', ',')
            
            if final_vos != '-':
                ws[vos_col + str(sec_row)] = final_vos
                ws[vos_col + str(sec_row + 1)] = str(round(vos + final_vos, 2)).replace('.', ',')
            else:
                ws[vos_col + str(sec_row)] = 0
                ws[vos_col + str(sec_row + 1)] = str(round(vos, 2)).replace('.', ',')

            # Fill head data    
            ws['A1'] = str(ws['A1'].value).replace('май', get_month(datetime.now().strftime("%d-%m-%Y")))
            ws['B3'] = date_from
            ws['C3'] = date_to
            ws['B4'] = datetime.now().strftime("%d-%m-%Y")
            ws['B5'] = order_type
            ws['B6'] = head_data['consumer']
            ws['B7'] = head_data['order']
            ws['B8'] = head_data['adress']
            ws['B11'] = head_data['cold_temp']
            ws['B12'] = head_data['factory_num']
            ws['B13'] = head_data['complex_num']

            curr_dir = self.save_dir + '/Output/' + head_data['save_folder']
            if not os.path.exists(curr_dir):
                os.makedirs(curr_dir)   
            template.save(curr_dir + '/' + file_name + '.xlsx')
            report += curr_dir + '/' + file_name + '.xlsx'+ '\n\n'

        return report