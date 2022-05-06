import pandas as pd
import numpy as np
import datetime
from datetime import date, timedelta

from simple_smartsheet import Smartsheet
smartsheet = Smartsheet('****')

#import gg master list
master_list=smartsheet.sheets.get(id=3713813712267140)
master_list=master_list.as_dataframe()

today = datetime.date.today()
wknm = datetime.date.today().isocalendar()[1]



currentweeknum = int(str(2022)+str(wknm))#202218
previousweeknum = int(str(2022)+str(wknm-1))

#read current and previous CO vs WO
file_name_c = 'CO info VS CPO and WO info '+ today.strftime("%m-%d-%Y") #'CO info VS CPO and WO info 05-05-2022'
file_name_p = f"{previousweeknum} Order status report"

#current week report file
current_wk=pd.read_excel(f'{file_name_c}.xls',sheet_name='CO info VS CPO and WO info',skiprows=2)

#previous week report file
previous_wk = pd.read_excel(f'{file_name_p}.xlsx',engine='openpyxl',sheet_name='CO info VS CPO and WO info')
orderAtRisk = pd.read_excel(f'{file_name_p}.xlsx',engine='openpyxl',sheet_name='Orders At Risk') # order at risk in the previous week

#drop some columns
cols = list(np.intersect1d(current_wk.columns.tolist(),previous_wk.columns.tolist()))
previous_wk = previous_wk[cols]

#identify changes from the previous week
chng = pd.merge(current_wk, previous_wk, how='inner', on='CO#', suffixes=('', '_previous'))
chng = chng[['CO#', 'Close','Shipment\nDate','Shipment\nDate_previous']]
n_chng = chng[(chng['Shipment\nDate'] > chng['Shipment\nDate_previous']) & (chng['Close']==0)].reset_index(drop=True)

#update the new df of the changes
new=pd.merge(current_wk,n_chng,how='left',on='CO#',suffixes=('','_y'))
new.drop(new.filter(like='_y').columns.tolist(),axis=1,inplace=True)

#add some imp't columns
new['O-week'] = (new['Original\nShip\nDate'].dt.strftime("%Y").astype('float')*100) + new['Original\nShip\nDate'].dt.strftime('%V').astype('float')
new['R-week'] = (new['Shipment\nDate'].dt.strftime("%Y").astype('float')*100) + new['Shipment\nDate'].dt.strftime('%V').astype('float')
new['R-week'] = [new['O-week'].get(i) if np.isnan(j) else new['R-week'].get(i) for i,j in enumerate(new['R-week'])]
new['MD week'] = (new['MD\nDate'].dt.strftime("%Y").astype('float')*100)+new['MD\nDate'].dt.strftime('%V').astype('float')
new['Delay'] = np.where(new['Original\nShip\nDate'] < new['Shipment\nDate'], 1, None)
new['Shipment Date Revised'] = np.where(np.isnan(new['Shipment\nDate_previous']), 'No', 'Yes')


#merge with gg master list
m_lst = pd.merge(new,master_list,how='left', on='Q#',suffixes=('','_y')).drop_duplicates(subset=['CO#'])
m_lst = m_lst[['Q#','Budget Customer', 'Customer Type', 'Pad Type', 'Molding SMV', 'Die Cut SMV', 'Machine Cut SMV', 'Ultrasonic SMV',
            'Hand Cut SMV', 'QC SMV', 'Packing SMV', 'Print SMV', 'QC+Packing SMV']]

new_m = pd.merge(new,m_lst,how='left', on='Q#',suffixes=('','_y')).drop_duplicates(subset=['CO#'])
new_m.drop(new_m.filter(like='_y').columns.tolist(),axis=1,inplace=True)
new_m = new_m.reset_index(drop=True)

a = new_m.select_dtypes(include=np.dtype('datetime64[ns]')).columns.tolist()

new_m.loc[:, a] = new_m.loc[:, a].apply(lambda x: x.dt.date)

#if E-Date not updated take M-Date
new_m['E-Date'] = new_m['E-Date'].fillna(0)
new_m['E-Date'] = np.where(new_m['E-Date'] == 0, new_m['MD\nDate'], new_m['E-Date'])
new_m['E-Date'] = new_m['E-Date'].replace(0, np.nan)


#find balance of each stage
new_m['CutBalance'] = new_m['Total\nCO Qty'].fillna(0)-new_m['Cut\nQty'].fillna(0)
new_m['MoldBalance'] = new_m['Total\nCO Qty'].fillna(0)-new_m['Mold\nQty'].fillna(0)
new_m['TrimBalance'] = new_m['Total\nCO Qty'].fillna(0)-new_m['Trim\nQty'].fillna(0)
new_m['QC+PackingBalance'] = new_m['Total\nCO Qty'].fillna(0) - new_m[['QC\nQty','PCK1\nQty']].min(axis=1,skipna=False)

#compile smv data from gg master list
# new_m[['Molding SMV','QC+Packing SMV']]=new_m[['Molding SMV','QC+Packing SMV']]
new_m['Trimming SMV']=new_m['Die Cut SMV'].fillna(0)+new_m['Machine Cut SMV'].fillna(0) +new_m['Ultrasonic SMV'].fillna(0) + new_m['Hand Cut SMV'].fillna(0)

#LTs
new_m['LT_IQC'] = 3
new_m['LT_Lam'] = 2
new_m['LT_Cut'] = 1
new_m['LT_FinalPack'] = 1

new_m['LT_Mold'] = np.ceil(new_m['MoldBalance'].fillna(0) * new_m['Molding SMV'].fillna(0) * (1/(0.8 * 2 * 60 * 10.5)))
new_m['LT_Trim'] = np.ceil(new_m['TrimBalance'].fillna(0) * new_m['Trimming SMV'].fillna(0) * (1/(0.7 * 2 * 60 * 10.5)))
new_m['LT_QCPack'] = np.ceil(new_m['QC+PackingBalance'].fillna(0) * new_m['QC+Packing SMV'].fillna(0) * (1/(0.7 * 2 * 60 * 10.5)))

#if status is not updated or blank change to '0'
new_m[['RM\nReceive\nSts','RM\nIssue\nSts']] = new_m[['RM\nReceive\nSts','RM\nIssue\nSts']].fillna(0,axis=1)

new_m['LT'] = new_m[['LT_Mold','LT_Trim','LT_QCPack']].max(axis=1,skipna=False) + np.where(((new_m['Cut\nQty']==0) & (new_m['RM\nIssue\nSts']==2)),new_m['LT_Cut'],0) + \
                np.where(((new_m['Cut\nQty']==0) & (new_m['RM\nIssue\nSts']!=2) & (new_m['RM\nReceive\nSts']==2)), new_m['LT_Lam'] + new_m['LT_Cut'],0) + \
                    np.where(new_m['RM\nReceive\nSts']!=2,new_m['LT_IQC'] + new_m['LT_Lam'] + new_m['LT_Cut'],0) + new_m['LT_FinalPack']


new_m['Exp_ShipDate'] = new_m[['RM\nReceive\nSts','LT','E-Date']].apply(lambda x: np.busday_offset(date.today(),offsets=x['LT'],roll='forward',weekmask=[1,1,1,1,1,1,0]) if x['RM\nReceive\nSts']==2 \
                            else (np.busday_offset(x['E-Date'],offsets=x['LT'],roll='forward',weekmask=[1,1,1,1,1,1,0])) if pd.notna(x['E-Date']) else None,axis=1)

new_m['Exp_ShipDate']=new_m['Exp_ShipDate'].dt.date
new_m['Risk']=np.where(new_m['Exp_ShipDate']>new_m['Shipment\nDate'],1,None)

new_m['Risk Type'] = np.where((new_m['Risk'] == 1) & (new_m['RM\nReceive\nSts'] == 2), 'LT', np.where((new_m['Risk'] == 1) & (new_m['RM\nReceive\nSts'] != 2),'Material Issue',''))

#pick columns to export to excel file
new_m = new_m[['WO#', 'Style', 'Colour', 'PO#', 'CO#', 'CO\nType', 'Q#', 'Customer', 'Brand','Budget Customer', 'Customer Type',
 'Pad Type', 'CO\nOpen\nDate', 'Original\nShip\nDate','O-week','Shipment\nDate','R-week','Delay',
 'Shipment Date Revised','Shipment\nDate_previous', 'Balance', 'MD\nDate',
 'MD week','Actual\nShip Date', 'Sales Qty\n(prs)', 'Sales\nAmount\n(HKD)', 'Close', 'Call Lot', 'Supplier', 'Branch',
 'Customer\nRequest\nDate', 'CO\nEx-Fty\nDate', 'E-Date', 'Packing\nList\nQty', 'Description', 'Product\nType',
 'Customer\nStyle', 'Total\nCO Qty', 'WO\nClose', 'Production\nStart Date', 'WO\nEx-Fty\nDate', 'WO\nQty',
 'RM\nReceive\nSts', 'RM\nInspection\nSts', 'RM\nIssue\nSts', 'Cut Pcs\nDate', 'Cut\nFty', 'Cut\nQty', 'Mold\nFty',
 'Mold\nSub-\nDept', 'Mold\nQty', 'Trim\nFty', 'Trim\nQty', 'QC\nFty', 'QC\nQty', 'PCK1\nFty', 'PCK1\nQty', 'PCK2\nFty',
 'PCK2\nQty', 'Exp\nFty', 'WO\nRemark', 'DyeLot\nRemark','Molding SMV', 'QC+Packing SMV','Trimming SMV', 'CutBalance',
 'MoldBalance', 'TrimBalance', 'QC+PackingBalance', 'LT_IQC', 'LT_Lam', 'LT_Cut', 'LT_FinalPack', 'LT_Mold', 'LT_Trim',
 'LT_QCPack', 'LT', 'Exp_ShipDate', 'Risk','Risk Type']]


#delayed orders
Delay = pd.DataFrame()
Delay = new_m[['WO#', 'CO#', 'Q#','Customer','Budget Customer','Customer Type', 'Original\nShip\nDate','Shipment\nDate',
             'Shipment Date Revised','Balance','Total\nCO Qty','Delay','CO\nType','Close']]

Delay = Delay[(Delay['CO\nType'] == 'FM') & (Delay['Delay'] == 1) & (Delay['Close'] == 0)]
Delay.sort_values('Shipment\nDate',inplace=True)
Delay.drop(['Delay','CO\nType','Close'],axis=1,inplace=True)
Delay.rename(columns={"Shipment Date Revised": "Revised Shipment Date changed from previous report"}, inplace=True)

#not updtated(ship date)
Not_updated_ship = pd.DataFrame()
Not_updated_ship=new_m[['WO#', 'CO#', 'Q#','Customer','Budget Customer','Customer Type','Shipment\nDate',
                         'Balance','Total\nCO Qty','CO\nType','Close']]
Not_updated_ship=Not_updated_ship[(Not_updated_ship['CO\nType'] == 'FM') & (Not_updated_ship['Close'] == 0) &
                                  (Not_updated_ship['Shipment\nDate'] < today-timedelta(days=2))]# 2days buffer for scan and system update(d-2)
Not_updated_ship.sort_values('Shipment\nDate',inplace=True)
Not_updated_ship.drop(['CO\nType','Close'],axis=1,inplace=True)

#not updtated(rm receive)
Not_updated_rm = pd.DataFrame()
Not_updated_rm =new_m[['WO#', 'CO#', 'Q#','Customer','Budget Customer','Customer Type','E-Date',
                         'Balance','Total\nCO Qty','CO\nType','Close','RM\nReceive\nSts']]
Not_updated_rm = Not_updated_rm[(Not_updated_rm['CO\nType'] == 'FM') & (Not_updated_rm['Close'] == 0) &
                                  (Not_updated_rm['E-Date'] < today) & (Not_updated_rm['RM\nReceive\nSts']!=2)]
Not_updated_rm.sort_values('E-Date',inplace=True)
Not_updated_rm.drop(['CO\nType','Close','RM\nReceive\nSts'],axis=1,inplace=True)

#orders of current week and the next week
This_wk=pd.DataFrame()
This_wk = new_m[['WO#', 'CO#', 'Q#','Customer','Budget Customer','Customer Type','Shipment\nDate','Balance','Total\nCO Qty',
               'Cut\nQty','Mold\nQty','Trim\nQty','QC\nQty','PCK1\nQty','PCK2\nQty','CO\nType','Close','R-week']]

This_wk = This_wk[(This_wk['CO\nType'] == 'FM') & (This_wk['Close'] == 0) & (This_wk['R-week'].isin([currentweeknum, (currentweeknum+1)]))]
This_wk.sort_values('Shipment\nDate',inplace=True)
This_wk.drop(['CO\nType','Close','R-week'],axis=1,inplace=True)


Risk = pd.DataFrame()
Risk = new_m[['WO#', 'CO#', 'Q#','Customer','Budget Customer','Customer Type','Shipment\nDate','Risk Type','Balance',
      'Total\nCO Qty','CO\nType','Close','Risk']]
Risk = Risk[(Risk['CO\nType'] == 'FM') & (Risk['Close'] == 0) & (Risk['Risk'] == 1)]
Risk.sort_values('Shipment\nDate',inplace=True)
Risk.drop(['CO\nType','Close','Risk'],axis=1,inplace=True)

shippedOrders = pd.read_excel('April 2022 GG shipped orders job costing list.xls',sheet_name='April 2022',skiprows=2)
orderAtRisk=pd.merge(orderAtRisk, shippedOrders, how='left', on='CO#', suffixes=('', '_previous'))
orderAtRisk[['Shipment\nDate','Latest Ship\nDate']] = orderAtRisk[['Shipment\nDate','Latest Ship\nDate']].apply(lambda x: x.dt.date)

orderAtRisk['OTD'] = np.where(orderAtRisk['Shipment\nDate'] >= orderAtRisk['Latest Ship\nDate'], "{:.0%}".format(1),
                      np.where(orderAtRisk['Shipment\nDate'] <= today-timedelta(days=2),"{:.0%}".format(0),''))
orderAtRisk=orderAtRisk[['WO#', 'CO#', 'Q#', 'Customer', 'Budget Customer', 'Customer Type', 'Shipment\nDate','Latest Ship\nDate','OTD', 'Risk Type', 'Balance', 'Total\nCO Qty']]


#write to excel file
writer = pd.ExcelWriter(f'{currentweeknum} Order status report.xlsx', engine='xlsxwriter')

new_m.to_excel(writer,sheet_name='CO info VS CPO and WO info', index=False,header=True)
Delay.to_excel(writer,sheet_name='Delay', index=False,header=True)
Not_updated_ship.to_excel(writer,sheet_name='Not Updated(ship date)', index=False,header=True)
Not_updated_rm.to_excel(writer,sheet_name='Not Updated(RM recieve sts)',index=False,header=True)
This_wk.to_excel(writer,sheet_name='This Week+1',index=False,header=True)
Risk.to_excel(writer,sheet_name='Orders At Risk',index=False,header=True)
orderAtRisk.to_excel(writer,sheet_name='LW Risk orders OTD',index=False,header=True)

writer.save()

