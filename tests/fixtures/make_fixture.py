"""Regenerate sample.pptx: python3 -m pip install python-pptx && python3 make_fixture.py"""
import struct, zlib, io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.oxml.ns import qn
from lxml import etree


def png(w=4, h=4, rgb=(255, 0, 0)):
    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))
    chunk = lambda t, d: struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


prs = Presentation()

s1 = prs.slides.add_slide(prs.slide_layouts[0])
s1.shapes.title.text = "Hello Deck"
s1.placeholders[1].text = "Subtitle here"
s1.notes_slide.notes_text_frame.text = "Speaker notes"

s2 = prs.slides.add_slide(prs.slide_layouts[1])
s2.shapes.title.text = "Bullets"
tf = s2.placeholders[1].text_frame
tf.text = "First point"
p = tf.add_paragraph()
p.level = 1
r = p.add_run()
r.text = "Bold red"
r.font.bold = True
r.font.size = Pt(20)
r.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
r2 = p.add_run()
r2.text = " italic"
r2.font.italic = True

s3 = prs.slides.add_slide(prs.slide_layouts[6])
tb = s3.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
tb.text_frame.text = "Box text"
tb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
rect = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5), Inches(1), Inches(2), Inches(1))
rect.fill.solid()
rect.fill.fore_color.rgb = RGBColor(0x00, 0xFF, 0x00)
rect.text_frame.text = "Green"
s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5), Inches(3), Inches(1), Inches(1))
s3.shapes.add_picture(io.BytesIO(png()), Inches(1), Inches(3), Inches(1), Inches(1))
tbl = s3.shapes.add_table(2, 2, Inches(1), Inches(5), Inches(4), Inches(1)).table
for i, t in enumerate(["A1", "B1", "A2", "B2"]):
    tbl.cell(i // 2, i % 2).text = t
grp = s3.shapes.add_group_shape()
grp.shapes.add_textbox(Inches(7), Inches(5), Inches(2), Inches(1)).text_frame.text = "Grouped"

tri = s3.shapes.build_freeform(Inches(7), Inches(1))
tri.add_line_segments([(Inches(8), Inches(2)), (Inches(7), Inches(2))], close=True)
tri = tri.convert_to_shape()
tri.name = "Freeform"
line = s3.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(8), Inches(3), Inches(7), Inches(4))  # flipH
line.name = "Line"
grad = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(6), Inches(2), Inches(1))
grad.name = "Gradient"
grad.fill.gradient()
grad.fill.gradient_angle = 0
stops = grad.fill.gradient_stops
stops[0].color.rgb = RGBColor(0xFF, 0x00, 0x00)
stops[1].color.rgb = RGBColor(0x00, 0x00, 0xFF)
alpha = etree.SubElement(stops[1]._gs.find(qn("a:srgbClr")), qn("a:alpha"))
alpha.set("val", "50000")

# Designer-style shapes: a full-slide rectangle filled with the slide background,
# a 90deg-rotated + flipped gradient, and a freeform whose points use formula guides.
bgrect = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(1), Inches(1))
bgrect.name = "BgFill"
bgrect._element.set("useBgFill", "1")

flip = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(3), Inches(6), Inches(2), Inches(1))
flip.name = "FlippedGradient"
flip.fill.gradient()
flip.fill.gradient_angle = 0
flip.fill.gradient_stops[0].color.rgb = RGBColor(0x00, 0x00, 0x00)
flip.fill.gradient_stops[1].color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
flip._element.spPr.find(qn("a:xfrm")).set("flipH", "1")
flip._element.spPr.find(qn("a:gradFill")).find(qn("a:lin")).set("ang", "3000000")

guided = s3.shapes.build_freeform(Inches(5), Inches(5))
guided.add_line_segments([(Inches(6), Inches(6)), (Inches(5), Inches(6))], close=True)
guided = guided.convert_to_shape()
guided.name = "GuidedFreeform"
geom = guided._element.spPr.find(qn("a:custGeom"))
gd_lst = geom.find(qn("a:gdLst"))
for name, fmla in [("gx", "*/ 1 w 2"), ("gy", "+- h 0 0")]:
    gd = etree.SubElement(gd_lst, qn("a:gd"))
    gd.set("name", name)
    gd.set("fmla", fmla)
pts = geom.findall(".//" + qn("a:pt"))
pts[1].set("x", "gx")  # half the width
pts[1].set("y", "gy")  # full height

prs.save("sample.pptx")

# --- SmartArt: python-pptx can't create it, so add the parts to the zip directly.
# Only the cached drawing (what PowerPoint last rendered) matters for display.
import zipfile, shutil

DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
frame = (
    f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="90" name="Diagram"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
    f'<p:xfrm><a:off x="1000000" y="2000000"/><a:ext cx="4000000" cy="2000000"/></p:xfrm>'
    f'<a:graphic><a:graphicData uri="{DGM}"><dgm:relIds xmlns:dgm="{DGM}" xmlns:r="{REL}" r:dm="rId90" r:lo="rId91" r:qs="rId92" r:cs="rId93"/></a:graphicData></a:graphic></p:graphicFrame>'
)
drawing = (
    f'<dsp:drawing xmlns:dsp="{DSP}" xmlns:a="{A}"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>'
    f'<dsp:sp modelId="{{1}}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>'
    f'<dsp:spPr><a:xfrm><a:off x="10000" y="20000"/><a:ext cx="1000000" cy="400000"/></a:xfrm><a:prstGeom prst="chevron"><a:avLst><a:gd name="adj" fmla="val 30000"/></a:avLst></a:prstGeom>'
    f'<a:solidFill><a:schemeClr val="accent2"/></a:solidFill></dsp:spPr>'
    f'<dsp:style><a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style>'
    f'<dsp:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2800" b="1"/><a:t>Step</a:t></a:r></a:p></dsp:txBody>'
    f'<dsp:txXfrm><a:off x="130000" y="20000"/><a:ext cx="760000" cy="400000"/></dsp:txXfrm></dsp:sp>'
    f'<dsp:sp modelId="{{2}}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>'
    f'<dsp:spPr><a:xfrm><a:off x="10000" y="420000"/><a:ext cx="900000" cy="800000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    f'<a:solidFill><a:schemeClr val="accent2"><a:tint val="40000"/><a:alpha val="90000"/></a:schemeClr></a:solidFill></dsp:spPr></dsp:sp>'
    f'</dsp:spTree></dsp:drawing>'
)
data = (
    f'<dgm:dataModel xmlns:dgm="{DGM}" xmlns:a="{A}"><dgm:ptLst/><dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram">'
    f'<dsp:dataModelExt xmlns:dsp="{DSP}" relId="rId94" minVer="{DGM}"/></a:ext></dgm:extLst></dgm:dataModel>'
)
parts = {
    "ppt/diagrams/data1.xml": ("diagramData", data, "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"),
    "ppt/diagrams/layout1.xml": ("diagramLayout", f'<dgm:layoutDef xmlns:dgm="{DGM}"/>', "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"),
    "ppt/diagrams/quickStyle1.xml": ("diagramQuickStyle", f'<dgm:styleDef xmlns:dgm="{DGM}"/>', "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"),
    "ppt/diagrams/colors1.xml": ("diagramColors", f'<dgm:colorsDef xmlns:dgm="{DGM}"/>', "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"),
    "ppt/diagrams/drawing1.xml": (None, drawing, "application/vnd.ms-office.drawingml.diagramDrawing+xml"),
}
rels = "".join(
    f'<Relationship Id="rId{90 + i}" Type="{REL}/{t}" Target="../diagrams/{path.split("/")[-1]}"/>'
    for i, (path, (t, _, _)) in enumerate(parts.items()) if t
) + f'<Relationship Id="rId94" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing1.xml"/>'

shutil.move("sample.pptx", "sample.tmp")
with zipfile.ZipFile("sample.tmp") as src, zipfile.ZipFile("sample.pptx", "w", zipfile.ZIP_DEFLATED) as dst:
    for item in src.infolist():
        body = src.read(item.filename)
        if item.filename == "ppt/slides/slide3.xml":
            body = body.decode().replace("</p:spTree>", frame + "</p:spTree>").encode()
        elif item.filename == "ppt/slides/_rels/slide3.xml.rels":
            body = body.decode().replace("</Relationships>", rels + "</Relationships>").encode()
        elif item.filename == "[Content_Types].xml":
            overrides = "".join(f'<Override PartName="/{p}" ContentType="{ct}"/>' for p, (_, _, ct) in parts.items())
            body = body.decode().replace("</Types>", overrides + "</Types>").encode()
        dst.writestr(item, body)
    for path, (_, xml, _) in parts.items():
        dst.writestr(path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml)
import os
os.remove("sample.tmp")
