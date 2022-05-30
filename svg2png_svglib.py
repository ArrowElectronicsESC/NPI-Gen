from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF, renderPM

def convertSVG2PNG(svgPath):
##    print("\n\nSVG Properties >> {}".format(svgPath))
    drawing = svg2rlg(svgPath + ".svg")
    renderPM.drawToFile(drawing, svgPath + ".png", fmt="PNG")
