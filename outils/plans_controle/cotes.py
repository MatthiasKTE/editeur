import fitz, re, sys, json
sys.path.insert(0,__import__('os').path.dirname(__import__('os').path.abspath(__file__)))
from extraire import lignes
RA={0.4,0.8,1.6,3.2,6.3,12.5}
EXCL=re.compile(r'(?i)echelle|scale|d[ée]tail|trous|holes|repere|tol\.|pi[eè]ce|piece|assembl|coupe|section|view|vue|^[A-Z]\d{0,2}$|^\d{4}/\d{4}$|^1$|^\d+x$|^\(?\d+x\)?$')
def num(s): return float(s.replace(',','.'))
def gen_lin(v):
    for lim,t in [(3,0.1),(6,0.1),(30,0.2),(120,0.3),(400,0.5),(1000,0.8),(2000,1.2),(4000,2)]:
        if v<=lim: return t
    return 3
def gen_ray(v):
    return 0.2 if v<=3 else 0.5 if v<=6 else 1.0
PETIT=re.compile(r'^\s*([+-])?\s*(\d+(?:[.,]\d+)?)?\s*$')
def parse(f):
    p,ls=lignes(f); W,H=p.rect.width,p.rect.height
    yt=H+10
    ZX=[]
    for t,r,d in ls:
        if t.startswith('NOTE:'): ZX.append(fitz.Rect(r.x0-5,r.y0-5,r.x0+440,r.y0+190))
        if t=='DATE': ZX.append(fitz.Rect(r.x0-40,r.y0-60,W,H))
        if t.startswith('Remark'): ZX.append(fitz.Rect(r.x0-5,r.y0-5,W,H))
    # zone tableau de cartouche/notes : on garde aussi les tableaux de trous (au-dessus de NOTE ou à gauche)
    items=[]
    for t,r,d in ls:
        if re.search(r'VALABLE|READY FOR|^CMD ',t): continue
        if any(Z.intersects(r) for Z in ZX): continue
        if r.y0>=yt-2 and not re.search(r'drilled|counter|M\d',t): 
            # tableaux de trous situés sous le dessin : on accepte les nombres purs à gauche du cartouche
            if not (r.x1 < W*0.45 and re.fullmatch(r'\d+(?:[.,]\d+)?',t) and r.y0<H-150): continue
        if r.x0>W-180 and r.y1<170: continue
        t2=t.replace('=','').strip()
        if len(re.findall(r'[A-Za-zÀ-ÿ]',t2))>12 and not re.search(r'(?i)drilled|counter|depth|max',t2): continue
        if '->' in t2: continue
        if not t2 or EXCL.search(t2) and not re.search(r'\d',t2.split()[0] if t2.split() else ''): continue
        if re.search(r'(?i)echelle|scale|d[ée]tail|trous|holes|repere|tol\.|pi[eè]ce|piece|assembl|coupe|section|view',t2): continue
        if re.fullmatch(r'(?![RM]\d)[A-Z]\d{0,2}|\d{4}/\d{4}|1|\d+x|\(\d+x\)',t2): continue
        items.append([t2,r,d])
    tables=[]
    for t,r,d in ls:
        if t=='REPERE':
            hx=[(tt,rr) for tt,rr,dd in ls if abs(rr.y0-r.y0)<4 and rr.x0>=r.x0-2 and rr.x0<r.x0+420]
            col={tt.split()[0]:rr for tt,rr in hx}
            labs=[(tt,rr) for tt,rr,dd in ls if re.fullmatch(r'[A-Z]\d{1,2}',tt) and abs((rr.x0+rr.x1)/2-(r.x0+r.x1)/2)<15 and rr.y0>r.y1]
            if not labs: continue
            ymax=max(rr.y1 for tt,rr in labs)+20
            xmax=col.get('Holes',r).x1+60
            reg=fitz.Rect(r.x0-5,r.y0-3,xmax,ymax)
            tc=col.get('Tol.')
            tolr=[tt for tt,rr,dd in ls if tc and reg.intersects(rr) and rr.x0>=tc.x0-5 and rr.x1<=tc.x1+5 and rr.y0>tc.y1 and re.search(r'\d[.,]\d+',tt)]
            tol=num(re.search(r'\d[.,]\d+',tolr[0]).group()) if tolr else None
            tables.append(dict(reg=reg,labs=labs,col=col,tol=tol))
    tab_out=[]
    for T in tables:
        for lab,lr in sorted(T['labs'],key=lambda x:(x[1].y0)):
            yc=(lr.y0+lr.y1)/2
            for ax in ('X','Y'):
                c=T['col'].get(ax)
                if not c: continue
                cand=[(tt,rr) for tt,rr,dd in ls if abs((rr.y0+rr.y1)/2-yc)<5 and abs((rr.x0+rr.x1)/2-(c.x0+c.x1)/2)<25 and re.fullmatch(r'\d+(?:[.,]\d+)?',tt)]
                if cand:
                    tt,rr=cand[0]; v=num(tt); tl=T['tol'] if T['tol'] is not None else gen_lin(v)
                    tab_out.append(dict(txt=f'{lab} {ax}={tt}',rect=list(rr),dir=[1,0],sym=f'Pos. {ax}',dim=v,sup=tl,inf=-tl,note=f'trou {lab}'))
        hc=T['col'].get('Holes')
        if hc:
            hl=[(tt,rr) for tt,rr,dd in ls if T['reg'].intersects(rr) and rr.x1>=hc.x0-40 and rr.y0>hc.y1]
            used=set()
            for tt,rr in sorted(hl,key=lambda x:(x[1].y0,x[1].x0)):
                if id(rr) in used: continue
                if re.fullmatch(r'M\d+',tt):
                    dep=[(t2,r2) for t2,r2 in hl if abs(r2.y0-rr.y0)<4 and r2.x0>rr.x1 and re.fullmatch(r'\d+',t2)]
                    tab_out.append(dict(txt=tt,rect=list(rr),dir=[1,0],sym='M',dim=tt,sup=None,inf=None,note='taraudage'+(f' prof. {dep[0][0]}' if dep else '')))
                elif re.search(r'drilled|counter',tt):
                    v=num(re.search(r'\d+(?:[.,]\d+)?',tt).group())
                    tab_out.append(dict(txt=tt,rect=list(rr),dir=[1,0],sym='Ø',dim=v,sup=gen_lin(v),inf=-gen_lin(v),note=re.sub(r'[\d.,]+','',tt).strip()))
                    if 'counter' in tt:
                        dep=[(t2,r2) for t2,r2 in hl if abs(r2.y0-rr.y0)<4 and r2.x0>rr.x1 and re.fullmatch(r'\d+',t2)]
                        if dep:
                            fr=[t2 for t2,r2 in hl if r2.x0>dep[0][1].x1-2 and abs(r2.y0-rr.y0)<12 and re.fullmatch(r'\d[.,]\d+',t2)]
                            sp=num(fr[0]) if fr else gen_lin(num(dep[0][0]))
                            tab_out.append(dict(txt='prof. '+dep[0][0],rect=list(dep[0][1]),dir=[1,0],sym='Prof.',dim=num(dep[0][0]),sup=sp,inf=0.0,note='profondeur lamage'))
    treg=[T['reg'] for T in tables]
    items=[it for it in items if not any(R.intersects(it[1]) for R in treg)]
    # séparer petits fragments de tolérance
    items=[it for it in items if not re.fullmatch(r'\d{1,2}/\d{1,2}/\d{4}',it[0])]
    def estfrag(t): return bool(PETIT.match(t)) and (t in('+','-','0') or t[:1] in '+-' or (re.fullmatch(r'\d[.,]\d+',t) and num(t)<1))
    mains=[it for it in items if not estfrag(it[0])]
    frag=[]
    for it in items:
        if not estfrag(it[0]): continue
        t,r,d=it
        z=fitz.Rect(r.x0-24,r.y0-16,r.x1+8,r.y1+16)
        if t in('+','-','0') or t[:1] in '+-' or any(z.intersects(mm[1]) and re.search(r'\d',mm[0]) for mm in mains): frag.append(it)
        else: mains.append(it)
    out=[]
    for t,r,d in mains:
        m=dict(txt=t,rect=list(r),dir=list(d),sym='',dim=None,sup=None,inf=None,note='')
        # fragments proches
        zone=fitz.Rect(r.x0-6,r.y0-14,r.x1+22,r.y1+14)
        att=[f for f in frag if zone.intersects(f[1]) and not f[3:]]
        for f in att: f.append('pris')
        sup=inf=None; sign=None; vals=[]
        for f in sorted(att,key=lambda f:(f[1].y0,f[1].x0)):
            mm=re.match(r'^\s*([+-])?\s*(\d+(?:[.,]\d+)?)?\s*$',f[0])
            if not mm: continue
            s,v=mm.group(1),mm.group(2)
            if s and v: (sup:=num(v)) if s=='+' else (inf:=-num(v))
            elif s and not v: sign=s
            elif v: vals.append(num(v))
        for v in vals:
            if v==0:
                if sup is None and inf is not None: sup=0.0
                elif inf is None: inf=0.0
            elif sign=='+' and sup is None: sup=v
            elif sign=='-' and inf is None: inf=-v
        tt=t
        g=re.search(r'±\s*(\d+(?:[.,]\d+)?)',tt)
        mmin=re.search(r'\s-\s*(\d[.,]\d+)',tt)
        if mmin: inf=-num(mmin.group(1)); tt=tt[:mmin.start()]
        fit=re.search(r'\b([A-Za-z]\d{1,2})\b',tt)
        if re.match(r'G\s*\d+/\d+',tt):
            m.update(sym='G',dim=tt,note='taraudage gaz')
        elif re.search(r'M\s*\d',tt):
            mt=re.search(r'(\d+)\s*x\s*',tt); md=re.search(r'M\s*(\d+(?:[.,]\d+)?)',tt)
            m.update(sym='M',dim='M'+md.group(1),note=(mt.group(1)+'x ' if mt else '')+'taraudage'+(' '+tt[md.end():].strip() if tt[md.end():].strip() else ''))
        elif re.search(r'(?i)drilled|counter',tt):
            v=num(re.search(r'\d+(?:[.,]\d+)?',tt).group()); m.update(sym='Ø',dim=v,note=re.sub(r'[\d.,]+','',tt).strip())
            if sup is None: sup,inf=gen_lin(v),-gen_lin(v)
        elif re.search(r'x\s*45|à\s*\d+°|x\s*\d+°',tt):
            m.update(sym='Chanfrein',dim=tt,note='')
        elif re.search(r'°',tt):
            m.update(sym='Angle',dim=tt)
        elif re.match(r'(?i)\(?\d*x?\)?\s*R\s*\d',tt) or re.search(r'(?i)^R\s*\d',tt):
            v=num(re.search(r'R\s*(\d+(?:[.,]\d+)?)',tt).group(1)); m.update(sym='R',dim=v,note=re.sub(r'R\s*[\d.,]+','',tt).strip())
            if 'max' in tt: sup,inf=0,-v
            elif sup is None: sup,inf=gen_ray(v),-gen_ray(v)
        else:
            nm=re.match(r'^(?:\(\d+x\)\s*|\d+\s*x\s*)?(?:depth\s*)?(\d+(?:[.,]\d+)?)',tt) or re.search(r'(\d+(?:[.,]\d+)?)',tt)
            if not nm: continue
            v=num(nm.group(1))
            # rugosité / tolérance géométrique
            if not g and sup is None and inf is None and re.fullmatch(r'\d[.,]\d+(\s*[A-Z])?',tt) and v<1 or (v in RA and r.height<15 and not g and sup is None and inf is None and re.fullmatch(r'\d[.,]\d',tt)):
                if v in RA and re.fullmatch(r'\d[.,]\d',tt) : continue      # rugosité Ra : non reprise
                m.update(sym='Géom.',dim=0,note='tolérance géométrique '+tt)
                sup,inf=v,0
            else:
                m.update(dim=v)
                if fit and fit.group(1)[0].isalpha(): m['sym']=fit.group(1)
                if 'depth' in tt: m['note']='profondeur'
                if re.search(r'\bØ|drilled',tt): m['sym']='Ø'
                if g: sup,inf=num(g.group(1)),-num(g.group(1))
                if sup is None and inf is None: sup,inf=gen_lin(v),-gen_lin(v)
                if sup is None: sup=0.0
                if inf is None: inf=0.0
        m['sup'],m['inf']=sup,inf
        out.append(m)
    # tri : haut→bas, gauche→droite (par bandes)
    out.sort(key=lambda m:(round(m['rect'][1]/60),m['rect'][0]))
    return p,out+tab_out
if __name__=="__main__":
    for f in sys.argv[1:]:
        p,o=parse(f); print("=====",f,len(o))
        for i,m in enumerate(o,1): print(i,m['sym'],m['dim'],m['sup'],m['inf'],'|',m['txt'],'|',m['note'])
