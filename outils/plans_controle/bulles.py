import fitz, sys, json, re
sys.path.insert(0,__import__('os').path.dirname(__import__('os').path.abspath(__file__)))
from cotes import parse
ROUGE=(0.85,0.05,0.05)
def bulle(page,rect,n,dirv):
    r=fitz.Rect(rect)
    R=7.5 if n<100 else 9
    if abs(dirv[1])>0.5:   # texte vertical : bulle à gauche
        cx,cy=r.x0-R-2, r.y0+R
    else:                  # bulle en haut à gauche
        cx,cy=r.x0-R-1, r.y0-R+3
    c=fitz.Point(cx,cy)
    page.draw_circle(c,R,color=ROUGE,fill=(1,1,1),width=0.9)
    s=str(n); fs=7.5 if n<100 else 6.5
    w=fitz.get_text_length(s,fontname='hebo',fontsize=fs)
    page.insert_text(fitz.Point(cx-w/2,cy+fs*0.36),s,fontname='hebo',fontsize=fs,color=ROUGE)
def traiter(src,dst,jout,over=None):
    p,items=parse(src)
    over=over or {}
    keep=[]
    for m in items:
        k=m['txt']
        if k in over.get('drop',[]): continue
        if k in over.get('set',{}): m.update(over['set'][k])
        keep.append(m)
    doc=fitz.open(src); page=doc[0]
    for i,m in enumerate(keep,1):
        bulle(page,m['rect'],i,m['dir']); m['n']=i
    doc.save(dst,garbage=3,deflate=True)
    json.dump(keep,open(jout,'w'),ensure_ascii=False,indent=0)
    return keep
