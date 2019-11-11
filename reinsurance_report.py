# -*- coding: utf-8 -*-
"""
Created on Mon Jul 29 08:53:41 2019

@author: martin.vavpotic
"""

import pandas as pd

Format = {'head' : {'font_name':'Arial', 'font_color':'black', 'bg_color':'yellow', 'font_size':'10', 'align':'left', 'valign':'bottom', 'text_wrap': True, 'bold': True, 'border': True},
   'spreadsheet' : {'font_name':'Arial', 'font_color':'black', 'bg_color': '#FFFFFF', 'font_size': '10', 'text_wrap': False, 'bold': False, 'border': True},
        'normal' : {'font_name':'Arial', 'font_color':'black', 'bg_color':'none', 'font_size':'10', 'type': 'General','text_wrap': False, 'bold': False, 'border': False},
          'bold' : {'font_name':'Arial', 'font_color':'black', 'bg_color':'none', 'font_size':'10', 'text_wrap': False, 'bold': True, 'border': False},
        'border' : {'font_name':'Arial', 'font_color':'black', 'bg_color':'none', 'font_size':'10', 'text_wrap': False, 'bold': False,'border': True}}
          
columns ={'MoscowRe':[['A:A', 8.43], ['B:B', 12.14], ['C:C',  8.43], ['D:D', 13.57], ['E:E', 12.29], ['F:G',  8.43], ['H:I', 13.57],
                      ['J:K', 8.43], ['L:N', 15.57], ['O:O',  8.43], ['P:P', 11.57], ['Q:Q', 12.57], ['R:R', 10.29], ['S:S', 12.57]],
       'StockholmRe':[["A:C", 8.43], ["D:D", 16.14], ["E:H",  8.43], ["I:I", 10.43], ["J:J",  8.43], ["K:K",  9.71], ["L:L",  8.43], ["M:P", 14.29]],
          'LondonRe':[["A:A", 8.43], ["B:B", 12.14], ["C:C",  8.43], ["D:E", 13.57], ["F:G",  8.43], ["H:I", 13.57], ["J:K",  8.43],
                      ["L:N",15.57], ["O:O",  8.43], ["P:Q", 12.57], ["R:R", 10.29], ["S:S", 12.57]],
         'HamburgRe':[["A:A", 6.29], ["B:C", 11.86], ["E:E", 19.29], ["F:G",  8.71], ["H:H", 12.86], ["I:J",  8.71],
                      ["K:K", 9.14], ["L:L", 11.29], ["M:O", 14.14], ["P:P", 12.71], ["Q:Q", 10.71], ["R:R", 11]]}
    
Risk_Reins = {'LondonRe':{'Term':'50001',
                       'Permanent Accidental Disability':'80',
                       'Accidental Death':'81',
                       'Accidental Annuity':'803'},
              'MoscowRe':{'Death':'50001',
                       'Broken Bone':'50006',
                       'Death&Crit Illness':'50007',
                       'Fractures,Dislocations':'50010',
                       'Mobility Aid':'50011',
                       'Rehabilitation':'50012',
                       'Accidental Disability':'80',
                       'Permanent Disability':'800',
                       'TPD':'804',
                       'Critical Illness':'805',
                       'Hospitalization':'806',
                       'Permanent disability 50%':'807',
                       'Accidental Death':'81',
                       'Short Term Disability':'84',
                       'Hospital Daily Allowance':'86'},
              'HamburgRe':{'Term':'50001',
                       'Critical Illness':'805',
                       'LTD':'804',
                       'Short Term Disability':'84',
                       'Accidental Death':'81',
                       'Permanent Accidental Disability':'800',
                       'Permanent Accidental Disability 25%':'808',
                       'Permanent Accidental Disability 50%':'807',
                       'Severe Fractures':'50006',
                       'Fractures Dislocation':'50010',
                       'Mobility Aid':'50011',
                       'Rehabilitation':'50012',
                       'Dislocation Benefit':'50013',
                       'Major Burns':'50014',
                       'Home Assistance':'50502',
                       'Hospital Cash':'50504',
                       'Home Care':'50505'},
            'StockholmRe':{}}
              
quarters_n={'period': {'year': ['01.01.', '31.12.'],
                         'Q1': ['01.01.', '31.03.'],
                         'Q2': ['01.04.', '30.06.'],
                         'Q3': ['01.07.', '30.09.'],
                         'Q4': ['01.10.', '31.12.']},
              'report':{'year':('31.01.', '28.02.', '31.03.', 
                                '30.04.', '31.05.', '30.06.', 
                                '31.07.', '31.08.', '30.09.', 
                                '31.10.', '30.11.', '31.12.'),
                          'Q1':('31.01.', '28.02.', '31.03.'),
                          'Q2':('30.04.', '31.05.', '30.06.'),
                          'Q3':('31.07.', '31.08.', '30.09.'),
                          'Q4':('31.10.', '30.11.', '31.12.')}}

def write_to_excel(excel_writer,frame,sheet_name):
    row_index = 0
    for i,spreadsheet in enumerate(dict(frame)[sheet_name]):
        if i == 0:
            dict(frame)[sheet_name][i].to_excel(excel_writer=excel_writer, sheet_name=sheet_name, index=False, startrow=row_index, startcol=0, header=False)
            row_index +=1
        else:
            dict(frame)[sheet_name][i].to_excel(excel_writer=excel_writer, sheet_name=sheet_name, index=False, startrow=row_index, startcol=0, header=True)
            row_index +=4
        row_index += dict(frame)[sheet_name][i].shape[0]
    #Mworksheet = excel_writer.sheets[sheet_name]
    #Mworksheet = set_format_m(Mworksheet, sheet_name, MoscowReFrame, 'MoscowRe')
    #return Mworksheet
        
def moscow_to_excel(ReinsArray,excel_writer):
    sheet_dict = dict(ReinsArray)
    for sheet_name in sheet_dict:
        write_to_excel(excel_writer=excel_writer, frame=ReinsArray, sheet_name=sheet_name)
        Mworksheet = excel_writer.sheets[sheet_name]
        Mworksheet = set_format_m(Mworksheet, sheet_name, ReinsArray, 'MoscowRe')

def hamburg_to_excel(excel_writer, ReinsArray, TitleArray, sheet_name):
    start_row = 0
    ReinsDict = dict(ReinsArray)
    for i,spreadsheet in enumerate(ReinsDict[sheet_name]):
        if i == 0:
            spreadsheet.to_excel(excel_writer=excel_writer, sheet_name=sheet_name, index=False, startrow=start_row, startcol=0, header=False)
        else:
            if 'Policies' in sheet_name:
                TitleArray[i].to_excel(excel_writer=excel_writer, sheet_name=sheet_name, index=False, startrow=start_row, startcol=0, header=False)
                start_row +=1
            spreadsheet.to_excel(excel_writer=excel_writer, sheet_name=sheet_name, index=False, startrow=start_row, startcol=0, header=True)
            start_row +=1
        start_row += ReinsDict[sheet_name][i].shape[0] +2 # Povečam vsakič za enako
            
def set_format(sheet, workbook, Reinsurer, data_1, data_2):
    sheet.set_row(8,64.5)
    for column in columns[Reinsurer]:
        sheet.set_column(column[0],column[1])
    sheet.conditional_format(data_1.shape[0] +1, 0, data_1.shape[0] +1, data_2.shape[1] -1,
                             {'type'  : 'no_blanks',
                              'format': workbook.add_format(Format['head'])})
    sheet.conditional_format(data_1.shape[0] +2, 0, data_1.shape[0] +1 +data_2.shape[0], data_2.shape[1] -1,
                             {'type'  : 'no_error',
                              'format': workbook.add_format(Format['spreadsheet'])})
    return sheet

def set_format_m(sheet, workbook, Worksheet, ReinsArray, Reinsurer):
    start_row = 0
    #sheet = writer.sheets[Worksheet]
    for column in columns[Reinsurer]:
        sheet.set_column(column[0],column[1])
    ReinsDict = dict(ReinsArray)
    for i,pair in enumerate(ReinsDict[Worksheet]):
    #for i in range(1,len(MoscowReFrame[Worksheet])):
    
        if i == 0: #First spreadsheet is informative; is not formatted, does add into start_row.
            start_row += pair.shape[0]
            if Reinsurer == 'HamburgRe':
                start_row += 2
                #if 'Policies' in Worksheet:
                    #start_row -= 1
        else:
            sheet.conditional_format(start_row +1, 0, start_row +1, ReinsDict[Worksheet][i].shape[1] -1,
                                     {'type'  : 'no_blanks',
                                      'format': workbook.add_format(Format['head'])})
            sheet.set_row(start_row +1,64.5)
            sheet.conditional_format(start_row +1 +1, 0, start_row +1 +pair.shape[0], pair.shape[1] -1,
                                     {'type'  : 'no_error',
                                      'format': workbook.add_format(Format['spreadsheet'])})
            start_row += pair.shape[0] +4
    return sheet
    

def selection_f(my_input, quarter, year, Reinsurer, Coverage=None, Product=None, Benefit=None, TimeDate='vel'):
    
    #Reinsurer is always specified, no need to handle exceptions.
    mask_reinsurer = (my_input['Reinsurer'] == Reinsurer)

    #Risk is specified everywhere except StockholmRe. Use try/except to avoid KeyError in case being used by StockholmRe
    #(no keys are specified for that Reinsurer, only one product).
    
    try:
        mask_risk = (my_input['COVERAGE_CODE'] == Risk_Reins[Reinsurer][Coverage])
    except KeyError:
        mask_risk = (my_input['Currency'] != None) #This is True everywhere, this is a neutral mask.

    #Product is specified in MoscowRe and HamburgRe.
    if Product == None:
        mask_product = (my_input['Currency'] != None) #This is True everywhere, this is a neutral mask.
    else:
        Product = Product.replace(" ", "").split(',')
        mask_product = (my_input['PRODUCT_NAME'].isin(Product))

    #Benefit is used in HamburgRe.
    if Benefit == None:
        mask_benefit = (my_input['Currency'] != None) #This is True everywhere, this is a neutral mask.
    else:
        Benefit = Benefit.replace(" ", "").split(',')
        mask_benefit = (my_input['BENEFIT'].isin(Benefit))

    #TimeDate is always specified.
    if TimeDate =='vel':
        mask_timedate = (my_input['POLICY_CANCEL_DATE'].isnull())
    elif TimeDate =='alt':
        mask_timedate = (my_input['POLICY_CANCEL_DATE'] >= pd.to_datetime(quarters_n['period'][quarter][0] + year)) & \
                        (my_input['POLICY_CANCEL_DATE'] <  pd.to_datetime(quarters_n['period'][quarter][1] + year))
                        
    #Default value = all True (no filter)
    mask = mask_reinsurer & mask_risk & mask_product & mask_benefit & mask_timedate

    my_output = my_input[mask]
    my_output = my_output.reset_index(drop=True)
    my_output.index += 1
    my_output = polish_data(my_output) # Polishes the formatting of dates and ages.
    my_output = my_output.reset_index(drop=False).rename(index=str, columns={"index": "N"}).drop(['Reinsurer'], axis =1)
    return my_output


def polish_data(data):
    data = convert_date_to_string(data)
    data = convert_age_to_int(data)
    return data


def convert_date_to_string(data):
    columns = ['REPORT_PER','POLICY_START_DATE','BENEFIT_START_DATE','POLICY_CANCEL_DATE','POLICY_MATURITY_DATE']
    for column in columns:
        data[column] = data[column].apply(lambda x: x.strftime('%d.%m.%Y') if not pd.isnull(x) else ' ')
    return data


def convert_age_to_int(data):
    columns = ['Curr Age']
    for column in columns:
        data[column] = data[column].apply(lambda x: int(x) if not pd.isnull(x) else ' ')
    return data    


def sql_time_string(quarter,year):
    sql_time = "and report_per in ("
    for each_date in quarters_n['report'][quarter]:
        sql_time += ("'" + each_date + year + "', ")
    sql_time = sql_time[0:-2]
    sql_time = sql_time + ') '
    return sql_time


def merge_combine(quarter, year, my_input):
    #First month is stored in select_comb by default. Others are added to it.
    select_comb = selection_time(my_input, quarters_n['report'][quarter][0]+year)
    
    # Start addition with second month, use compare_months method to do it.
    for each_date in quarters_n['report'][quarter][1:]:
        select_new = selection_time(my_input, each_date + year)
        select_comb = compare_months(select_comb, select_new)
        
    return select_comb


def compare_months(select_1, select_2):
    select_merged = select_1.merge(select_2, on=['Currency','POLICY_NUMBER','OSA_ID','INS1_SEX','POLICY_START_DATE','BENEFIT_START_DATE','POLICY_MATURITY_DATE', \
                                                 'PRODUCT_NAME','COVERAGE_CODE', 'BENEFIT','Reinsurer'], how='outer', indicator=True)

    mask_left = (select_merged['_merge'] == 'left_only')
    mask_both = (select_merged['_merge'] == 'both')
    mask_right = (select_merged['_merge'] == 'right_only')

    # Handling premium in case it changes due to change in premium rate. In left and right mask, take one that is not NaN. In 'both', it's summed up.
    select_merged.loc[mask_left, 'Premium w Load'] = select_merged['Premium w Load_x']
    select_merged.loc[mask_right, 'Premium w Load'] = select_merged['Premium w Load_y']
    select_merged.loc[mask_both, 'Premium w Load'] = select_merged['Premium w Load_x'] + select_merged['Premium w Load_y']

    select_left = select_merged[['REPORT_PER_x','Currency','POLICY_NUMBER','OSA_ID','Curr Age_x','INS1_SEX','POLICY_START_DATE','BENEFIT_START_DATE', \
                                 'POLICY_CANCEL_DATE_x', 'POLICY_MATURITY_DATE', 'TARIFA_x','SUM_INSURED_x','Sum_at_risk_x','Rate_x','Premium w Load', \
                                 'Reinsurance Ratio_x','PRODUCT_NAME','COVERAGE_CODE', 'BENEFIT','Retention Limit_x','Reinsurer', '_merge']][mask_left]
    select_left = select_left.rename(columns={'REPORT_PER_x':'REPORT_PER','Curr Age_x':'Curr Age','POLICY_CANCEL_DATE_x':'POLICY_CANCEL_DATE','TARIFA_x':'TARIFA', \
                                              'SUM_INSURED_x':'SUM_INSURED', 'Sum_at_risk_x':'Sum_at_risk','Rate_x':'Rate','Reinsurance Ratio_x':'Reinsurance Ratio', \
                                              'Retention Limit_x':'Retention Limit'})

    select_rt_bt = select_merged[['REPORT_PER_y','Currency','POLICY_NUMBER','OSA_ID','Curr Age_y','INS1_SEX','POLICY_START_DATE','BENEFIT_START_DATE', \
                                  'POLICY_CANCEL_DATE_y','POLICY_MATURITY_DATE', 'TARIFA_y','SUM_INSURED_y','Sum_at_risk_y','Rate_y','Premium w Load', \
                                  'Reinsurance Ratio_y','PRODUCT_NAME','COVERAGE_CODE','BENEFIT','Retention Limit_y', 'Reinsurer', '_merge']][~mask_left]
    select_rt_bt = select_rt_bt.rename(columns={'REPORT_PER_y':'REPORT_PER','Curr Age_y':'Curr Age','POLICY_CANCEL_DATE_y':'POLICY_CANCEL_DATE','TARIFA_y':'TARIFA', \
                                                'SUM_INSURED_y':'SUM_INSURED', 'Sum_at_risk_y':'Sum_at_risk','Rate_y':'Rate','Reinsurance Ratio_y':'Reinsurance Ratio', \
                                                'Retention Limit_y':'Retention Limit'})
       
    select_comb = pd.concat([select_left,select_rt_bt])
    select_comb = select_comb.drop(columns='_merge', axis=1)
    return select_comb


def selection_time(my_input, mesec):
    mask_time = (my_input['REPORT_PER'] == mesec)
    my_output = my_input[mask_time]
    my_output = my_output.reset_index(drop=True)
    my_output.index += 1
    my_output.sort_values(by=['POLICY_NUMBER','OSA_ID','BENEFIT','COVERAGE_CODE'])
    
    return my_output


def filter_non_reinsured(data):
    mask_1 = data['Sum_at_risk'] == 0
    
    return data.loc[~mask_1]


def filter_period(data, quarter, year):
    mask_1 = data['POLICY_CANCEL_DATE'] < pd.to_datetime(quarters_n['period'][quarter][0] + year)
    
    return data.loc[~mask_1]