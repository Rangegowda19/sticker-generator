"""
export_cdr_data.py — writes a data file the CorelDRAW VBA macro reads to draw
native, editable stickers (a real .cdr). Python does ALL the layout math and
outputs exact positions in INCHES, so the VBA side is dumb-simple.

Format: a plain text file, one line per drawing element:
  PAGE|<n>                                  start of a new page
  TEXT|x|y|size_pt|hscale|anchor|string      a text element (x,y in points, top-left origin)
  OVAL|cx|cy|rx|ry|linewidth                 an ellipse (points)
Coordinates: origin top-left of an A4 page, y downward, in POINTS (1/72").
CorelDRAW VBA converts these to document units.
"""
import os
import common
import layout_eight, layout_four, layout_two, layout_one
from reportlab.pdfbase.pdfmetrics import stringWidth

PT = 72.0
PW, PH = 595.276, 841.89
LAYOUTS = {8: layout_eight, 4: layout_four, 2: layout_two, 1: layout_one}

def _sw(txt, size):
    return stringWidth(txt or "", common.FONT, size)

def _positions(lay, x, y, w, h, d, show_dest):
    def g(name, default=None): return getattr(lay, name, default)
    has_dest = bool(show_dest and d.get("dest"))
    keys = (["dest"] if has_dest else []) + ["item","code","vol","desc"]
    if has_dest:
        LINE_H=g("LINE_H_P3",g("LINE_H")); LINE_GAP=g("LINE_GAP_P3",g("LINE_GAP",None)); su=g("SHIFT_UP_P3",g("SHIFT_UP",0))
    else:
        LINE_H=g("LINE_H_P2",g("LINE_H")); LINE_GAP=g("LINE_GAP_P2",g("LINE_GAP",None)); su=g("SHIFT_UP_P2",g("SHIFT_UP",0))
    F={k:v*PT for k,v in LINE_H.items()}; W={k:v*PT for k,v in lay.LINE_W.items()}
    if g("MATCH_5LINE_SPACING",False) and not has_dest:
        F={k:v*PT for k,v in g("LINE_H_P3",lay.LINE_H).items()}
        LINE_GAP=g("LINE_GAP_P3",g("LINE_GAP",None)); layout_keys=["dest","item","code","vol","desc"]
    else:
        layout_keys=keys
    padV=h*lay.PAD_V_FRACTION; total=sum(F[k] for k in layout_keys)
    gap=(h-2*padV-total)/max(1,len(layout_keys)-1) if LINE_GAP is None else LINE_GAP*PT
    vol_adj = g("VOL_GAP_ADJUST", 0) * PT   # +inch moves vol (and desc) UP
    desc_adj = g("DESC_GAP_ADJUST", 0) * PT # +inch moves desc UP

    # EXACT production baseline gaps (reverse-engineered) override even spacing.
    bg = g("BASELINE_GAPS_P3") if has_dest else g("BASELINE_GAPS_P2")
    base={}
    if bg:
        # place each line using the measured gap from the previous baseline
        yy=y+h-padV-F[layout_keys[0]]+su*PT
        pair_order=["dest_item","item_code","code_vol","vol_desc"]
        base[layout_keys[0]]=yy
        for i in range(1,len(layout_keys)):
            step = bg.get(pair_order[i-1], None)
            if step is None:
                yy -= gap+F[layout_keys[i]]
            else:
                yy -= step*PT
            base[layout_keys[i]]=yy
    else:
        yy=y+h-padV-F[layout_keys[0]]+su*PT
        for i,k in enumerate(layout_keys):
            base[k]=yy
            if i+1<len(layout_keys): yy-=gap+F[layout_keys[i+1]]
        # independent gap tweaks (positive = move that line up / tighter gap above it)
        if "vol" in base:  base["vol"]  += vol_adj
        if "desc" in base: base["desc"] += vol_adj + desc_adj
    # FONT sizes (points): production-measured, separate from line height
    fp = g("FONT_PT_P3") if has_dest else g("FONT_PT_P2")
    if fp is None:
        Ffont = dict(F)
    else:
        Ffont = {k: fp.get(k, F.get(k)) for k in ["dest","item","code","vol","desc"]}
    return F,W,base,keys,Ffont

def _scale_for(lay,W,key,txt,size):
    if key in W and txt:
        nat=_sw(txt,size)
        if nat>0: return max(20,min(160,W[key]/nat*100))
    return common.TEXT_HSCALE

def _emit_sticker(out, lay, x, y, w, h, d, show_dest):
    F,W,base,keys,Ff=_positions(lay,x,y,w,h,d,show_dest)
    sr=getattr(lay,"SHIFT_RIGHT",0); cx=x+w/2+sr*PT
    def td(yb): return PH-yb   # bottom-up -> top-down points

    if "dest" in keys:
        f=Ff["dest"]; sc=_scale_for(lay,W,"dest",d["dest"],f)
        out.append(f'TEXT|{cx:.2f}|{td(base["dest"]):.2f}|{f:.2f}|{sc:.2f}|middle|{d["dest"]}')

    item=d["item_no"] or ""; f1=Ff["item"]; sc1=_scale_for(lay,W,"item",item,f1)
    out.append(f'TEXT|{cx:.2f}|{td(base["item"]):.2f}|{f1:.2f}|{sc1:.2f}|middle|{item}')
    item_right=cx+_sw(item,f1)*sc1/100/2

    f2=Ff["code"]; sc2=_scale_for(lay,W,"code",d["code"] or "",f2)
    out.append(f'TEXT|{item_right:.2f}|{td(base["code"]):.2f}|{f2:.2f}|{sc2:.2f}|end|{d["code"] or ""}')

    vsep=getattr(lay,"VOL_SEP",common.VOL_SEP)
    vol=f"({d['volume'].strip('()') if d['volume'] else ''})"
    pcs=f"{d['pcs']} PCS" if d["pcs"] else ""; L=common.CIRCLE_LETTER
    mid=f"{vol}{vsep}{L}{vsep}{pcs}"; f3=Ff["vol"]; sc3=_scale_for(lay,W,"vol",vol,f3)
    # cap the full volume-line width to the block width, so the whole text block
    # matches production width (keeps vol/desc from stretching wider than item)
    bw_in = getattr(lay, "BLOCK_WIDTH_IN", None)
    if bw_in:
        full_w = _sw(mid, f3) * sc3/100
        if full_w > bw_in*PT:
            sc3 = sc3 * (bw_in*PT) / full_w
    out.append(f'TEXT|{item_right:.2f}|{td(base["vol"]):.2f}|{f3:.2f}|{sc3:.2f}|end|{mid}')
    total_w=_sw(mid,f3)*sc3/100; left_x=item_right-total_w
    vw=_sw(vol+vsep,f3)*sc3/100; d_center=left_x+vw+_sw(L,f3)*sc3/100/2
    # oval sized from LETTER point size (production letter slightly bigger than vol)
    letter_pt=getattr(lay,"LETTER_PT",f3)
    orx=getattr(lay,"OVAL_RX",common.OVAL_RX)*letter_pt; ory=getattr(lay,"OVAL_RY",common.OVAL_RY)*letter_pt
    olw=getattr(lay,"OVAL_LINE_WIDTH",common.OVAL_LINE_WIDTH)
    ocy=td(base["vol"]+f3*0.35)
    out.append(f'OVAL|{d_center:.2f}|{ocy:.2f}|{orx:.2f}|{ory:.2f}|{olw:.2f}')

    desc=d["description"] or ""; f4=Ff["desc"]; natural=_sw(desc,f4)
    maxw=item_right-left_x; scale=min(common.TEXT_HSCALE,(maxw/natural)*100 if natural else 100)
    midx=(left_x+item_right)/2
    out.append(f'TEXT|{midx:.2f}|{td(base["desc"]):.2f}|{f4:.2f}|{scale:.2f}|middle|{desc}')

def export(designs, out_path, sheets_for_design, copies_for_height, show_dest=False):
    lines=["FONT|Arial","PAGEW|%.2f"%PW,"PAGEH|%.2f"%PH]
    page=0
    for d in designs:
        common.CIRCLE_LETTER=d.get("circle_letter") or common.CIRCLE_LETTER
        n,stickers,sheets=sheets_for_design(d)
        lay=LAYOUTS.get(n,layout_one)
        cols,rows=lay.GRID; sw=lay.STICKER_W_IN*PT; sh=lay.STICKER_H_IN*PT
        grid_w=cols*sw; y_top=PH
        x0=(PW-grid_w)/2 if lay.PLACEMENT=="center" else PW-grid_w-2*2.83
        for _ in range(sheets):
            page+=1; lines.append(f"PAGE|{page}")
            i=0
            for r in range(rows):
                for c in range(cols):
                    x=x0+c*sw; y=y_top-(r+1)*sh
                    _emit_sticker(lines,lay,x,y,sw,sh,d,show_dest)
                    i+=1
                    if i>=n: break
    # Windows line endings so CorelDRAW VBA Line Input reads each line
    with open(out_path,"w",encoding="utf-8",newline="") as f:
        f.write("\r\n".join(lines))
    return page