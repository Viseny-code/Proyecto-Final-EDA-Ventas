from pathlib import Path
import csv,json
import matplotlib;matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches,Pt,RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
P=Path(__file__).resolve().parent;s=json.loads((P/'resultados.json').read_text())
def rows(n):
 with open(P/(n+'.csv'),encoding='utf-8-sig') as f:return list(csv.DictReader(f))
def plot(x,y,title,path,rotation=0):
 fig,ax=plt.subplots(figsize=(9.4,4));ax.bar(x,y,color='#0B837C');ax.set_title(title,loc='left',fontweight='bold');ax.grid(axis='y',alpha=.2);ax.set_axisbelow(True)
 ax.tick_params(axis='x',labelrotation=rotation);fig.tight_layout();fig.savefig(P/path,dpi=165,bbox_inches='tight');plt.close(fig)
annual=rows('anio');products=sorted(rows('producto'),key=lambda x:float(x['beneficio']),reverse=True);regions=sorted(rows('region'),key=lambda x:float(x['beneficio']),reverse=True)
plot([r['year'] for r in annual],[float(r['ventas'])/1e9 for r in annual],'Ventas por año · miles de millones','evolucion.png')
plot([r['producto'] for r in products[:7]],[float(r['beneficio'])/1e9 for r in products[:7]],'Top 7 productos por beneficio · miles de millones','productos.png',35)
plot([r['region'] for r in regions],[float(r['beneficio'])/1e9 for r in regions],'Beneficio por región · miles de millones','regiones.png',30)
doc=Document();sec=doc.sections[0];sec.header.paragraphs[0].text='PROYECTO FINAL  |  BUSINESS ANALYTICS';sec.top_margin=Inches(.83);sec.bottom_margin=Inches(.76)
styles=doc.styles;styles['Normal'].font.name='Aptos';styles['Normal'].font.size=Pt(10);styles['Normal'].paragraph_format.space_after=Pt(7)
for sty in ['Title','Heading 1','Heading 2']:styles[sty].font.name='Aptos Display';styles[sty].font.color.rgb=RGBColor(23,50,77)
doc.add_heading('EDA y dashboard de ventas internacionales',0);a=doc.add_paragraph('Análisis comercial, rentabilidad y eficiencia logística · 2020–2026');a.style='Subtitle'
doc.add_paragraph('Fuente: Proyecto_final_60000_pedidos_24_columnas(1).xlsx | Fecha de corte del análisis: 30/10/2026')
doc.add_heading('Resumen ejecutivo',1)
doc.add_paragraph(f"Se han auditado {s['filas']:,} pedidos y {s['columnas']} variables. Las ventas agregadas ascienden a {s['total_sales']/1e9:,.2f} miles de millones de unidades monetarias y el beneficio a {s['total_profit']/1e9:,.2f} miles de millones, con un margen ponderado del {s['margin_weighted_pct']:.2f} %. El tiempo medio pedido-envío es de {s['average_delivery_days']:.1f} días. Las cifras son propias de un conjunto SIMULADO; deben interpretarse como ejercicio de análisis y no como operaciones empresariales verificadas.")
doc.add_heading('1. Objetivos y preguntas de negocio',1)
for t in ['Cuantificar volumen, ventas, costes y rentabilidad y su evolución temporal.','Comparar productos, regiones, prioridades y canales de venta.','Analizar los plazos de expedición y su relación con los márgenes.','Construir un cuadro de mando operativo con filtros por año y región.','Identificar riesgos de calidad y proponer actuaciones medibles.']:doc.add_paragraph(t,style='List Bullet')
doc.add_heading('2. Procedencia, alcance y metodología',1)
doc.add_paragraph('El archivo contiene 60.000 registros elaborados sintéticamente a partir de una pestaña con 951 pedidos de origen. Se han tratado como registros operativos para los cálculos, sin cambiar ni ocultar su naturaleza simulada. El periodo de fecha de pedido comprende el 01/01/2020 y el 24/10/2026. La moneda no consta acreditada: se usan unidades monetarias del archivo. 2026 es parcial.')
doc.add_paragraph('Proceso reproducible en Python: lectura OOXML del Excel, verificación de tipos, validación de claves y fechas, recalculo de identidades económicas, comprobaciones de coherencia, agregaciones por dimensiones y contrastes inferenciales (SciPy). Se preserva la fuente original y se crean tablas agregadas independientes para el dashboard.')
doc.add_heading('3. Limpieza profunda y control de calidad',1)
doc.add_paragraph('Se revisaron los 24 campos por vacíos, la unicidad de ID Pedido, unidades y precios no negativos, igualdad ventas = unidades × precio, costes = unidades × coste, beneficio = ventas − costes, diferencia fecha envío − fecha pedido, año/mes extraídos y porcentajes de margen. En estas reglas no se encontraron errores ni valores ausentes. No se aplicaron imputaciones ni se eliminaron registros por no existir justificación para hacerlo.')
doc.add_paragraph('Se identificaron anomalías semánticas: al menos 307 pedidos de Afghanistan y 459 de Iran constan en la región África. Conviene auditar la asignación país-región del catálogo completo; no se ha reasignado automáticamente porque podría modificar la lógica de negocio original. Las categorías de entrega se mantienen como etiquetas del archivo; no se ha supuesto ningún acuerdo SLA.')
doc.add_heading('4. Análisis descriptivo',1)
tab=doc.add_table(rows=0,cols=2);tab.style='Light Shading Accent 1'
for k,v in [('Pedidos',f"{s['filas']:,}"),('Ventas',f"{s['total_sales']:,.2f}"),('Costes',f"{s['total_cost']:,.2f}"),('Beneficio',f"{s['total_profit']:,.2f}"),('Margen ponderado',f"{s['margin_weighted_pct']:.2f}%"),('Ticket medio',f"{s['average_ticket']:,.2f}"),('Ticket mediano',f"{s['median_order_sales']:,.2f}"),('Envío medio',f"{s['average_delivery_days']:.2f} días"),('Envío mediano',f"{s['median_delivery_days']:.0f} días")]:
 c=tab.add_row().cells;c[0].text=k;c[1].text=v
doc.add_paragraph('El ticket muestra elevada dispersión: percentil 25 = {:,.2f}, mediana = {:,.2f} y percentil 75 = {:,.2f}. El beneficio global no debe confundirse con la media simple de márgenes porcentuales.'.format(*s['quantiles_sales'][1:4]))
doc.add_picture(str(P/'evolucion.png'),width=Inches(6.3))
doc.add_paragraph('La serie de ventas se muestra para describir el volumen anual. No debe interpretarse el crecimiento como tendencia orgánica, dado que los pedidos fueron generados y 2026 es incompleto.')
doc.add_picture(str(P/'productos.png'),width=Inches(6.3));doc.add_picture(str(P/'regiones.png'),width=Inches(6.3))
doc.add_paragraph(f"Líder en beneficio: {products[0]['producto']} ({float(products[0]['beneficio'])/1e9:.2f} miles de millones). La región con mayor beneficio agregado es {regions[0]['region']} ({float(regions[0]['beneficio'])/1e9:.2f} miles de millones); el tamaño del segmento influye en esa clasificación.")
doc.add_heading('5. Análisis estadístico e hipótesis',1)
a=s['channel_margin_test'];doc.add_paragraph(f"H₀: el margen medio por pedido es igual en Offline y Online. Prueba t de Welch, bilateral: t = {a['t']:.3f}, p = {a['p']:.4f}. Márgenes medios: {a['channels'][0]} {a['means_pct'][0]:.2f}% y {a['channels'][1]} {a['means_pct'][1]:.2f}%. Con α = 0,05, no se rechaza H₀; ello no demuestra igualdad exacta.")
a=s['spearman_delivery_margin'];doc.add_paragraph(f"Relación entre días hasta envío y margen del pedido: Spearman ρ = {a['rho']:.4f}, p = {a['p']:.4g}. Existe evidencia estadística de una asociación de magnitud prácticamente nula; la significación se ve favorecida por el gran tamaño muestral, no debe presentarse como relación comercial relevante ni causal.")
a=s['kruskal_delivery_by_region'];doc.add_paragraph(f"Diferencias de plazos entre regiones: Kruskal-Wallis H = {a['h']:.3f}, p = {a['p']:.4f}. No se observa evidencia suficiente de distribuciones distintas de plazos por región con α = 0,05.")
ci=s['mean_order_profit_ci95'];doc.add_paragraph(f"Intervalo t al 95% para el beneficio medio por pedido: [{ci[0]:,.2f}, {ci[1]:,.2f}] unidades monetarias. Este intervalo presume un mecanismo de muestreo independiente y es meramente ilustrativo: en datos sintéticos no representa incertidumbre respecto de una población empresarial real.")
doc.add_heading('6. Dashboard operativo en Excel',1)
doc.add_paragraph('En Dashboard_Interactivo.xlsx se ofrece una hoja DASHBOARD con filtros desplegables de año y región (celdas B4 y D4). Las cifras y tablas de canal, producto y mes se calculan con fórmulas SUMPRODUCT sobre cubos preparados en Python; los gráficos se vinculan a estas tablas. Las pestañas BASE_AGREGADA, PRODUCTO_REGION y MES_REGION conservan los resúmenes auditables; LEEME explica su uso.')
doc.add_paragraph('KPIs: número de pedidos, ventas, beneficio y margen ponderado. Las dimensiones permiten contrastar año, región, canal, familia de producto y estacionalidad mensual. Las tablas corresponden a fecha del pedido.')
doc.add_heading('7. Conclusiones y recomendaciones',1)
for t in ['Priorizar el análisis de rentabilidad y ticket por producto y región, evitando seleccionar inversiones únicamente por volumen acumulado.','Investigar el diccionario de país-región y corregir el mapeo en una copia gobernada antes de usar el dashboard para decisiones geográficas.','Evitar afirmar diferencias de margen entre canales o plazos entre regiones: los contrastes no lo respaldan.','Definir SLA y estudiar métricas de cumplimiento de envío por prioridad y familia de producto; un promedio de días no equivale a entrega al cliente.','Comparar periodos equivalentes para evaluar 2026 y trabajar con datos operativos reales antes de emitir recomendaciones económicas definitivas.']:doc.add_paragraph(t,style='List Bullet')
doc.add_heading('8. Limitaciones y reproducibilidad',1)
doc.add_paragraph('No se dispone de moneda validada, devoluciones, descuentos, cliente real, costes de transporte ni fecha de recepción. El beneficio es contable según el dataset, no necesariamente beneficio neto. Los valores atípicos de ventas pueden ser legítimos, y no se han winsorizado ni borrado. La generación artificial puede inducir patrones no causales. El script analisis.py produce resultados.json y nueve tablas CSV a partir del archivo fuente; crear_dashboard.py reconstruye el panel y crear_informe.py reconstruye este informe.')
doc.add_paragraph('Anexo técnico: reglas de calidad verificadas, diccionario implícito de 24 campos en resultados.json, estadísticas, tablas agregadas y código Python entregado en el paquete de proyecto.')
footer=sec.footer.paragraphs[0];footer.text='EDA / Business Analytics · Documento académico · Datos simulados';footer.alignment=WD_ALIGN_PARAGRAPH.CENTER
out=P/'Informe_EDA_Proyecto_Final.docx';doc.save(out);print(out)
