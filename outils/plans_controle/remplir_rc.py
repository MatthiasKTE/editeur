import openpyxl, json, sys, copy, re, glob, os
from openpyxl.worksheet.cell_range import CellRange
rc, dossier = sys.argv[1], sys.argv[2]   # RC xlsx, dossier des cotes_<ref>.json
wb=openpyxl.load_workbook(rc)
def fnum(v):
    return v
for ws in wb.worksheets:
    a5=str(ws['A5'].value or '')
    m=re.search(r'Réf\. ([\d-]+)',a5)
    if not m: print('réf introuvable',ws.title); continue
    ref=m.group(1); f=os.path.join(dossier,f'cotes_{ref}.json')
    C=json.load(open(f))
    # pied de page : repérer la ligne « Contrôlé par »
    pied=next(r for r in range(9,ws.max_row+1) if str(ws.cell(r,1).value or '').startswith('Contr'))
    dispo=pied-9
    besoin=len(C)+1
    if besoin>dispo:
        n=besoin-dispo
        merges=[CellRange(str(x)) for x in ws.merged_cells.ranges if CellRange(str(x)).min_row>=pied]
        for x in merges: ws.unmerge_cells(str(x))
        ht={r:ws.row_dimensions[r].height for r in range(pied,ws.max_row+1)}
        ws.insert_rows(pied,n)
        for r,h in ht.items(): ws.row_dimensions[r+n].height=h
        for x in merges: ws.merge_cells(start_row=x.min_row+n,start_column=x.min_col,end_row=x.max_row+n,end_column=x.max_col)
        for r in range(pied,pied+n):
            ws.row_dimensions[r].height=ws.row_dimensions[10].height
            for c in range(1,20):
                ws.cell(r,c)._style=copy.copy(ws.cell(10,c)._style)
        pied+=n
    for r in range(9,pied):
        for c in range(1,20): ws.cell(r,c).value=None
    for i,m in enumerate(C):
        r=9+i
        ws.cell(r,1).value=m['n']
        sym=m['sym'] or ''
        note=m.get('note','')
        court={'M':'Filet.','G':'Filet.','Chanfrein':'Chanf.','Angle':'Angle','Géom.':'Géom.','Prof.':'Prof.','R':'R','Ø':'Ø'}
        if sym.startswith('Pos.'): b=sym[-1]+' '+note.replace('trou ','')
        else: b=court.get(sym,sym)
        ws.cell(r,2).value=b or None
        if isinstance(m['dim'],(int,float)) and m['sup'] is not None:
            ws.cell(r,3).value=m['dim']; ws.cell(r,4).value=m['sup']; ws.cell(r,5).value=m['inf']
            ws.cell(r,6).value=f'=C{r}+D{r}'; ws.cell(r,7).value=f'=C{r}+E{r}'
            for c in (3,4,5,6,7): ws.cell(r,c).number_format='0.0##'
        else:
            ws.cell(r,3).value=m['txt'] if sym not in ('M','G') else (m['txt']+(' – '+note.split('taraudage')[-1].strip() if 'prof' in note else ''))
        if sym=='Géom.': ws.cell(r,3).value=m['txt']; ws.cell(r,4).value=m['sup']; ws.cell(r,5).value=None; ws.cell(r,6).value=None; ws.cell(r,7).value=None
    aff=str(ws['A4'].value).split()[2]
    ws['A4']=str(ws['A4'].value)+f"   –   Plan de contrôle : PC_{aff}_{ref}.pdf"
    ws.print_title_rows='7:8'
    ws.page_setup.fitToHeight=0; ws.page_setup.fitToWidth=1
    print(ws.title, ref, len(C),'cotes')
wb.save(rc)
