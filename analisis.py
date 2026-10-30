"""EDA reproducible: lee XLSX mediante OOXML, valida, agrega, contrasta y exporta resultados."""
import zipfile,xml.etree.ElementTree as ET,collections,datetime as dt,json,math,statistics,csv
from pathlib import Path
import numpy as np
from scipy import stats
ROOT=Path(__file__).resolve().parent
SRC=ROOT/'archivo_original.xlsx'
if not SRC.exists():
    raise FileNotFoundError(f'Falta archivo_original.xlsx en la carpeta del proyecto: {ROOT}')
N='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
def read_sheet(zip_file,filename):
    with zip_file.open(filename) as f:
        for ev,el in ET.iterparse(f,events=('end',)):
            if el.tag!=N+'row':continue
            r={}
            for c in el:
                if c.tag!=N+'c':continue
                v=c.find(N+'v'); ins=c.find(N+'is'); value=v.text if v is not None else (''.join(x.text or '' for x in ins.iter(N+'t')) if ins is not None else None)
                if value is None:continue
                letters=''.join(x for x in c.get('r') if x.isalpha())
                idx=0
                for x in letters:idx=idx*26+ord(x)-64
                r[idx-1]=value
            yield r
with zipfile.ZipFile(SRC) as z:
    stream=read_sheet(z,'xl/worksheets/sheet1.xml');hdr=next(stream);cols=[hdr.get(i,'') for i in range(24)]
    data=[[r.get(i) for i in range(24)] for r in stream]
print('Leídos',len(data),'registros')
F=lambda x:float(x) if x not in (None,'') else float('nan')
records=[];issues=collections.Counter();ids=set(); geo=collections.Counter();cells_missing=collections.Counter(); category_delay=collections.Counter(); category_counts=collections.Counter()
for i,r in enumerate(data,2):
    d=dict(zip(cols,r))
    for k,v in d.items():
        if v is None or v=='':cells_missing[k]+=1
    try:
        units=F(d['Unidades']); price=F(d['Precio Unitario']);cost=F(d['Coste unitario']);sale=F(d['Importe venta total']);totalcost=F(d['Importe Coste total']);profit=F(d['BENEFICIO']);days=F(d['DIAS']); margin=F(d['Margen beneficio %']);benefit_unit=F(d['Beneficio por unidad'])
        order_date=dt.datetime(1899,12,30)+dt.timedelta(days=F(d['Fecha pedido']))
        ship_date=dt.datetime(1899,12,30)+dt.timedelta(days=F(d['Fecha envío']))
        if d['ID Pedido'] in ids:issues['ID pedido duplicado']+=1
        ids.add(d['ID Pedido'])
        if any(not math.isfinite(x) for x in (units,price,cost,sale,totalcost,profit,days)):issues['Valores numéricos no finitos']+=1
        if units<=0 or price<0 or cost<0:issues['Unidades/precios/costes inválidos']+=1
        if abs(sale-units*price)>0.02:issues['Venta ≠ unidades × precio']+=1
        if abs(totalcost-units*cost)>0.02:issues['Coste ≠ unidades × coste unitario']+=1
        if abs(profit-(sale-totalcost))>0.02:issues['Beneficio ≠ ventas − costes']+=1
        if abs((ship_date-order_date).days-days)>0.1:issues['Días de entrega incoherentes']+=1
        if days<0:issues['Envío anterior al pedido']+=1
        if int(F(d['Año pedido']))!=order_date.year or int(F(d['Mes pedido']))!=order_date.month:issues['Año/mes derivado incorrecto']+=1
        if abs(margin-(100*profit/sale if sale else 0))>0.06:issues['Margen porcentual incoherente']+=1
        if abs(benefit_unit-profit/units)>0.02:issues['Beneficio unitario incoherente']+=1
        if ship_date.date()>dt.date(2026,10,30):issues['Fecha de envío posterior a fecha de corte']+=1
        category_counts[d['Categoría de entrega']]+=1
        category_delay[(d['Categoría de entrega'],int(days))]+=1
        records.append({'id':d['ID Pedido'],'cliente':d['ID Cliente'],'region':d['Zona'],'pais':d['País'],'producto':d['Tipo de producto'],'canal':d['Canal de venta'],'prioridad':d['Prioridad'],'fecha':order_date.date().isoformat(),'year':order_date.year,'month':order_date.month,'quarter':(order_date.month-1)//3+1,'units':units,'sales':sale,'cost':totalcost,'profit':profit,'days':days,'margin':profit/sale if sale else np.nan})
    except Exception as ex:issues['Error de parseo']+=1
for r in records:
    if (r['pais']=='Afghanistan' and r['region']=='África') or (r['pais']=='Iran' and r['region']=='África'):
        geo[(r['pais'],r['region'])]+=1

def aggregate(keys):
    result={}
    for r in records:
        key=tuple(r[k] for k in keys);acc=result.setdefault(key,[0,0,0,0,0,0])
        acc[0]+=1;acc[1]+=r['sales'];acc[2]+=r['cost'];acc[3]+=r['profit'];acc[4]+=r['units'];acc[5]+=r['days']
    return [[*k,int(v[0]),*map(lambda x:round(x,2),v[1:])] for k,v in sorted(result.items())]
for name,keys in [('anio',['year']),('mes',['year','month']),('region',['region']),('producto',['producto']),('canal',['canal']),('prioridad',['prioridad']),('region_canal_anio',['year','region','canal']),('producto_anio',['year','producto']),('region_anio',['year','region']),('producto_region_anio',['year','region','producto']),('mes_region_anio',['year','region','month'])]:
    rows=aggregate(keys)
    with open(ROOT/(name+'.csv'),'w',newline='',encoding='utf-8-sig') as f:
        w=csv.writer(f);w.writerow([*keys,'pedidos','ventas','costes','beneficio','unidades','dias_acumulados']);w.writerows(rows)
arr=lambda col:np.array([r[col] for r in records],dtype=float)
sales=arr('sales');profit=arr('profit');delay=arr('days');margins=arr('margin');years=collections.Counter(r['year'] for r in records)
chan={key:np.array([r['margin'] for r in records if r['canal']==key]) for key in sorted(set(r['canal'] for r in records))}
channel_names=list(chan)
test=stats.ttest_ind(chan[channel_names[0]],chan[channel_names[1]],equal_var=False) if len(chan)==2 else None
rho,p_spear=stats.spearmanr(delay,margins)
regions=sorted(set(r['region'] for r in records));kr=stats.kruskal(*[np.array([r['days'] for r in records if r['region']==region]) for region in regions]);
np.random.seed(42);boots=np.mean(np.random.choice(profit,(1000,500),replace=True),axis=1) # representative bootstrap subsamples; CI prefer analytic below
ci=stats.t.interval(0.95,df=len(profit)-1,loc=np.mean(profit),scale=stats.sem(profit))
summary={'archivo':str(SRC.name),'filas':len(records),'columnas':len(cols),'columnas_nombres':cols,'registros_originales':951,'origen':'Datos simulados según pestaña Información','total_sales':round(float(sales.sum()),2),'total_profit':round(float(profit.sum()),2),'total_cost':round(float(arr('cost').sum()),2),'total_units':int(arr('units').sum()),'margin_weighted_pct':round(float(profit.sum()/sales.sum()*100),3),'average_ticket':round(float(sales.mean()),2),'average_delivery_days':round(float(delay.mean()),3),'median_delivery_days':float(np.median(delay)),'median_order_sales':float(np.median(sales)),'std_order_sales':round(float(np.std(sales,ddof=1)),2),'quantiles_sales':np.percentile(sales,[0,25,50,75,95,99,100]).round(2).tolist(),'annual_counts':dict(sorted(years.items())),'missing_by_column':dict(cells_missing),'issues':dict(issues),'geo_examples':dict((str(k),v) for k,v in geo.items()),'delivery_categories':dict(category_counts),'delivery_category_days':dict((str(k),v) for k,v in category_delay.items()),'channel_margin_test':{'channels':channel_names,'means_pct':[round(float(np.mean(chan[x])*100),3) for x in channel_names],'t':float(test.statistic) if test else None,'p':float(test.pvalue) if test else None},'spearman_delivery_margin':{'rho':float(rho),'p':float(p_spear)},'kruskal_delivery_by_region':{'h':float(kr.statistic),'p':float(kr.pvalue)},'mean_order_profit_ci95':list(map(float,ci)),'regions':regions,'categories':sorted(set(r['producto'] for r in records)),'channels':channel_names,'min_date':min(r['fecha'] for r in records),'max_date':max(r['fecha'] for r in records)}
(ROOT/'resultados.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps({k:summary[k] for k in ['filas','total_sales','total_profit','margin_weighted_pct','issues','annual_counts','channel_margin_test','spearman_delivery_margin','geo_examples']},ensure_ascii=False,indent=2))
