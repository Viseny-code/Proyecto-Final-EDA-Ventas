import csv,json
from pathlib import Path
from artifact_tool import Workbook,SpreadsheetFile
P=Path(__file__).resolve().parent; s=json.loads((P/'resultados.json').read_text())
w=Workbook.create(); dash=w.worksheets.add('DASHBOARD'); guide=w.worksheets.add('LEEME');
source=w.worksheets.add('BASE_AGREGADA'); prod=w.worksheets.add('PRODUCTO_REGION'); months=w.worksheets.add('MES_REGION')
def getcsv(nm):
 with (P/(nm+'.csv')).open(encoding='utf-8-sig',newline='') as f:return list(csv.reader(f))
def load(sh,name):
 rows=getcsv(name)
 vals=[]
 for ir,row in enumerate(rows):
  vals.append([int(x) if i in ([0,1,2,3] if False else []) else (float(x) if ir and (x.replace('.','',1).replace('-','',1).isdigit() and i>=len(rows[0])-6) else (int(x) if ir and rows[0][i] in ('year','month','pedidos') else x)) for i,x in enumerate(row)])
 sh.get_range_by_indexes(0,0,len(vals),len(vals[0])).values=vals
 sh.get_range_by_indexes(0,0,1,len(vals[0])).format={'fill':'#17324D','font':{'bold':True,'color':'#FFFFFF'},'row_height':29}
 sh.freeze_panes.freeze_rows(1)
 for j in range(len(vals[0])):sh.get_range_by_indexes(0,j,len(vals),1).format.column_width=18
 return len(vals)
nbase=load(source,'region_canal_anio');nprod=load(prod,'producto_region_anio');nmonth=load(months,'mes_region_anio')
navy='#17324D';teal='#0B837C';grey='#EAF1F6';white='#FFFFFF'
dash.merge_cells('A1:N2');dash.get_range('A1').values=[['PANEL EJECUTIVO  |  PEDIDOS Y RENTABILIDAD 2020–2026']];dash.get_range('A1:N2').format={'fill':navy,'font':{'color':white,'bold':True,'size':17},'row_height':27}
dash.get_range('A4').values=[['AÑO']];dash.get_range('C4').values=[['REGIÓN']];dash.get_range('B4').values=[['Todos']];dash.get_range('D4').values=[['Todas']];
for a in ['A4','C4']:dash.get_range(a).format={'font':{'bold':True,'color':navy}}
dash.get_range('B4').data_validation={'rule':{'type':'list','values':['Todos']+[str(y) for y in range(2020,2027)]}}
dash.get_range('D4').data_validation={'rule':{'type':'list','values':['Todas']+s['regions']}}
for a in ['B4','D4']:dash.get_range(a).format={'fill':'#DDF3EF','font':{'bold':True,'color':navy},'row_height':28}
# Use four SUMIFS branches instead of SUMPRODUCT/IF/VALUE to prevent
# filters from generating #VALUE! when the selector reads "Todos".
def agg_source(sheet, last, value_col, extras=()):
    amount=f"{sheet}!${value_col}$2:${value_col}${last}"
    years=f"{sheet}!$A$2:$A${last}"
    regions=f"{sheet}!$B$2:$B${last}"
    def sumif(with_year,with_region):
        args=[amount]
        if with_year: args += [years,'$B$4']
        if with_region: args += [regions,'$D$4']
        for col,cell in extras: args += [f'{sheet}!${col}$2:${col}${last}',cell]
        return f"SUMIFS({','.join(args)})" if len(args)>1 else f"SUM({amount})"
    return ('=IF($B$4="Todos",IF($D$4="Todas",'+sumif(False,False)+','+sumif(False,True)+'),'
            +'IF($D$4="Todas",'+sumif(True,False)+','+sumif(True,True)+'))')
def agg(col):return agg_source('BASE_AGREGADA',nbase,col)
kpis=[('B7','PEDIDOS','D'),('E7','VENTAS','E'),('H7','BENEFICIO','G'),('K7','MARGEN','')]
for a,label,column in kpis:
 col=a[0];dash.get_range(f'{col}6').values=[[label]];dash.get_range(f'{col}6').format={'font':{'color':teal,'bold':True,'size':11}}
 if column:dash.get_range(a).formulas=[[agg(column)]]
dash.get_range('K7').formulas=[['=IFERROR(H7/E7,0)']]
for a in ['B7','E7','H7','K7']:dash.get_range(a).format={'font':{'size':18,'bold':True,'color':navy},'fill':grey,'row_height':36}
dash.get_range('B7').set_number_format('#,##0');dash.get_range('E7').set_number_format('#,##0,," M"');dash.get_range('H7').set_number_format('#,##0,," M"');dash.get_range('K7').set_number_format('0.0%')
dash.get_range('A10:D10').values=[['CANAL','PEDIDOS','VENTAS','BENEFICIO']]
for i,ch in enumerate(s['channels'],11):
 dash.get_range(f'A{i}').values=[[ch]]
 for j,col in enumerate(['D','E','G'],1):
  dest='BCD'[j-1]
  dash.get_range(f'{dest}{i}').formulas=[[agg_source('BASE_AGREGADA',nbase,col,[('C',f'$A{i}')])]]
dash.get_range('F10:I10').values=[['PRODUCTO','PEDIDOS','VENTAS','BENEFICIO']]
# other source columns A year B region C product D count E sales F cost G profit
# Product breakdown uses the same selectors
for i,cat in enumerate(s['categories'],11):
 dash.get_range(f'F{i}').values=[[cat]]
 for cell,src in [('G','D'),('H','E'),('I','G')]:dash.get_range(f'{cell}{i}').formulas=[[agg_source('PRODUCTO_REGION',nprod,src,[('C',f'$F{i}')])]]
dash.get_range('A17:D17').values=[['MES','PEDIDOS','VENTAS','BENEFICIO']]
# Monthly breakdown uses the same selectors
for i in range(18,30):
 dash.get_range(f'A{i}').values=[[i-17]]
 for cell,src in [('B','D'),('C','E'),('D','G')]:dash.get_range(f'{cell}{i}').formulas=[[agg_source('MES_REGION',nmonth,src,[('C',f'$A{i}')])]]
for rng in ['A10:D10','F10:I10','A17:D17'] :dash.get_range(rng).format={'fill':navy,'font':{'color':white,'bold':True},'row_height':25}
for c in 'ABCDEFGHIJKLMN':dash.get_range(f'{c}:{c}').format.column_width=15
dash.get_range('F:F').format.column_width=23;dash.get_range('D:D').format.column_width=20
for rng in ['C11:D12','H11:I25','C18:D29']:dash.get_range(rng).set_number_format('#,##0,," M"')
# Charts on area L / lower left
for typ,range_,title,start,end in [('bar','A10:B12','Pedidos por canal','K10','N23'),('bar',f'F10:I{10+len(s["categories"])}','Rendimiento por producto','K25','N44'),('line','A17:C29','Evolución mensual','A32','I49')]:
 try:
  chart=dash.charts.add(typ,dash.get_range(range_));chart.title_text=title;chart.set_position(start,end)
 except Exception as exc:print('chart warn',title,str(exc)[:150])
dash.freeze_panes.freeze_rows(4)
notes=[['GUÍA DEL DASHBOARD'],['1. Selecciona AÑO en B4 (Todos o un año concreto).'],['2. Selecciona REGIÓN en D4 (Todas o una región).'],['3. Los indicadores, tablas y gráficos se recalculan con los filtros.'],['4. El valor económico se conserva en las unidades del archivo original; no hay divisa acreditada.'],['5. Los datos fueron simulados: resultados válidos para un ejercicio, no evidencia de ventas reales.'],['6. Los datos fuente completos se mantienen en el XLSX original, sin sobrescribir.'],['7. BASE_AGREGADA y otras pestañas son cubos reproducibles calculados en Python.'],['8. 2026 es un periodo parcial: no comparar su total directamente con años completos.'],['9. Los filtros se aplican a pedidos por año de PEDIDO, no año de envío.']]
guide.get_range_by_indexes(0,0,len(notes),1).values=notes;guide.get_range('A:A').format.column_width=112;guide.get_range('A1').format={'fill':navy,'font':{'bold':True,'color':white,'size':16},'row_height':36};guide.get_range('A2:A10').format.row_height=28
SpreadsheetFile.export_xlsx(w).save(str(P/'Dashboard_Interactivo.xlsx'))
print('OK dashboard',nbase,nprod,nmonth)
