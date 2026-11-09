import fitz, re, json, sys
def lignes(f):
    d=fitz.open(f); p=d[0]; W,H=p.rect.width,p.rect.height
    out=[]
    for b in p.get_text("dict")["blocks"]:
        for l in b.get("lines",[]):
            t=" ".join(s["text"].strip() for s in l["spans"] if s["text"].strip())
            if not t: continue
            out.append((t, fitz.Rect(l["bbox"]), l["dir"]))
    return p, out
if __name__=="__main__":
    p,ls=lignes(sys.argv[1])
    for t,r,d in ls:
        print(f"{r.x0:6.0f}{r.y0:6.0f}{r.x1:6.0f}{r.y1:6.0f} {d[0]:+.0f},{d[1]:+.0f} | {t}")
