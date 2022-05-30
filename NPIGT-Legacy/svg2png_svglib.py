from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF, renderPM
##import cairosvg


def convertSVG2PNG(svgPath, pngPath, scale=10):
    drawing = svg2rlg(svgPath)
    renderPM.drawToFile(drawing, pngPath, fmt="PNG")

    print("\n\nSVG Properties >> {}", svgPath)
    print(drawing.getProperties());
##    cairosvg.svg2png(url=svgPath, write_to=pngPath, scale=scale)
