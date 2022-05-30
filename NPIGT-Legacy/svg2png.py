import cairosvg


def convertSVG2PNG(svgPath, pngPath, scale=10):
    cairosvg.svg2png(url=svgPath, write_to=pngPath, scale=scale)
