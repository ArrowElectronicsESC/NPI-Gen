from pptx import Presentation
from pptx.util import Inches, Pt
import pptx.enum.shapes as pptx_types

from svg2png_svglib import convertSVG2PNG

import openpyxl

MAX_ROWS = 6;
global_AUX = None

def get_column_letter(column_index):
    """Return the column letter for a column index.
    For example, column_index 3 would return 'C'.
    """
    column_letter = ""
    while column_index > 0:
        column_letter = chr(ord('A') + (column_index-1) % 26) + column_letter
        column_index = int((column_index-1) / 26)
    return column_letter

def getExcelData(workbook, sheetName):
    pxl_doc = openpyxl.load_workbook(workbook)
    sheet = pxl_doc[sheetName]
    excelData = {}
    # iterate over columns in the first row
    for i in range(1, sheet.max_column+1):
        # get column letter
        col = get_column_letter(i)
        excelData[sheet.cell(row=1, column=i).value] = sheet.cell(row=2, column=i).value

    return excelData


def deleteImage(image):
    image = image._element;
    image.getparent().remove(image);
    
def deleteShape(shape):
    sp = shape._sp
    sp.getparent().remove(sp)

def getShapeProperties(shape):
    x = shape.left;
    y = shape.top;
    h = shape.height;
    w = shape.width;

    return x, y, h, w;

def fitImage(h_n, w_n, h, w):
    scale = h/h_n;
    h_n = h_n*scale;
    w_n = w_n*scale;

    if w_n > w:
        scale = w/w_n;
        h_n = h_n*scale;
        w_n = w_n*scale;

    return int(h_n), int(w_n);

def putImage(slide, shape, source, fmt):
    global global_AUX
    x, y, h, w = getShapeProperties(shape)

    img = source + "." + fmt;

    slide.shapes.add_picture(img, x, y);

    if source.find('logo') > 0:
        x, y, h, w = getShapeProperties(slide.shapes[-1]);
        
        logo = source.split('-');
        logo = logo[0];
        logo_file = open("{}-width.txt".format(logo));
        global_AUX = logo_file;
        logo_width = float(logo_file.read());

        scale = logo_width/w.inches;
        
        slide.shapes[-1].width = int(slide.shapes[-1].width*scale);
        slide.shapes[-1].height = int(slide.shapes[-1].height*scale);

        x, y, h, w = getShapeProperties(slide.shapes[-1]);
        
        slide.shapes[-1].top = int(y - h/2);

    else:
        slide.shapes[-1].height, slide.shapes[-1].width = fitImage(slide.shapes[-1].height, slide.shapes[-1].width, h, w);

    image = slide.shapes[-1];
##        scale = h/slide.shapes[-1].height;
##        slide.shapes[-1].width = int(slide.shapes[-1].width*scale);
##        slide.shapes[-1].height = int(slide.shapes[-1].height*scale);
    
    slide.shapes._spTree.insert(2, slide.shapes[-1]._element);

    deleteShape(shape);

    return image;

def iter_cells(table):
    for row in table.rows:
        for cell in row.cells:
            yield cell

def formatOPNTable(slide):
    global global_AUX
    for index, cell in enumerate(iter_cells(slide.shapes[-1].table)):
        for paragraph in cell.text_frame.paragraphs:
            for run in paragraph.runs:
                if index < 3:
                    run.font.size = Pt(16)
                else:
                    run.font.size = Pt(12)
                run.font.name = "Arrow Display"
                
def drawOPNTable(slide, shape, OPN):
    global MAX_ROWS
    x, y, h, w = getShapeProperties(shape)
    OPNs = OPN.split("/");

    if len(OPNs) > MAX_ROWS:
        print("More than <{}> P/Ns were identified, taking the first <{}> elements to preserve the shape".format(MAX_ROWS,MAX_ROWS))
        OPNs = OPNs[0:MAX_ROWS];
    
    slide.shapes.add_table(rows = len(OPNs) + 1,
                           cols = 3,
                           left = x,
                           top = y,
                           width = w,
                           height = int(len(OPNs) + 1*h/MAX_ROWS));

    writeOPNTableHeader(slide);

    deleteShape(shape);

def writeOPNTableHeader(slide):
    slide.shapes[-1].table.cell(0, 0).text = "OPN";
    slide.shapes[-1].table.cell(0, 1).text = "Description";
    slide.shapes[-1].table.cell(0, 2).text = "Package";
    

def writeOPNTable(slide, col, data_in):
    data = data_in.split("/");

    if len(data) > MAX_ROWS:
        data = data[0:MAX_ROWS];

    for i, val in enumerate(data):
        slide.shapes[-1].table.cell(i + 1, col).text = val;

def putData(slide, shape, name, data):
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            print(run.text)
            aux = run.text;

            if name.find("Image") > 0:
                if name.find("ogo") > 0:
                    sourceImage = "images//logo//{}".format(data[name][0:data[name].rfind(".")]);
                else:
                    sourceImage = "images//{}".format(data[name][0:data[name].rfind(".")]);
                    
                if data[name].find("svg") > 0:
                    convertSVG2PNG(sourceImage);
                    shape = putImage(slide,
                                     shape,
                                     sourceImage,
                                     "png");

                elif data[name].find("png") > 0:
                    shape = putImage(slide,
                                     shape,
                                     sourceImage,
                                     "png");

                elif data[name].find("jpg") > 0:
                    shape = putImage(slide,
                                 shape,
                                 sourceImage,
                                 "jpg");
                    
                else:
                    print("No format for <{}> detected".format(name));

            elif name.find("Table") > 0:
                drawOPNTable(slide,
                             shape,
                             data['OPNTable']);

                writeOPNTable(slide, 0, data['OPNTable']);
                writeOPNTable(slide, 1, data['OPNTableDescription']);
                writeOPNTable(slide, 2, data['OPNTablePackage']);
                formatOPNTable(slide)


            elif name.find("Link") >= 0:
##                        aux = run.text;
                run.text = "{} - {}".format(data["Supplier"], data["PartNumber"]);
                run.hyperlink.address = data[name];

            elif name.find("Text") >= 0:
##                        aux = run.text;
                run.text = data[aux];

            else:
##                        aux = run.text;
##                        print(getShapeProperties(dict_template[prs.slides.index(slide)][aux]['SHAPE']['object']))
                run.text = data[aux];
                shape.text_frame.fit_text(font_file = "ArrowDisplay_Md.ttf", max_size = 54);

                for run in shape.text_frame.paragraphs[0].runs:
                    run.font.name = 'Arrow Display';
##                        print(getShapeProperties(dict_template[prs.slides.index(slide)][aux]['SHAPE']['object']))
    x, y, h, w = getShapeProperties(shape);               
    return x, y, h, w, shape;






data = getExcelData("./NPI_TEMPLATE_FILL_Test.xlsx", 'TEMPLATE_1_FILL')

prs = Presentation("Template 99.pptx")

# text_runs will be populated with a list of strings,
# one for each text run in presentation
text_runs = []
dict_slide = {};
dict_template = {};
for slide in prs.slides:
    print(prs.slides.index(slide))
    if not prs.slides.index(slide) in dict_slide:
        dict_template[prs.slides.index(slide)] = {};
        dict_template[prs.slides.index(slide)]['object'] = slide;
        dict_template[prs.slides.index(slide)]['info'] = {};
        
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                print(run.text)
                aux = run.text;
                aux = aux.split('-');

                if aux[0] in data:
                    if not aux[0] in dict_template[prs.slides.index(slide)]['info']:
                        print("Creating index <{}>".format(aux[0]))
                        dict_template[prs.slides.index(slide)]['info'][aux[0]] = {};
                        if len(aux) > 1:
                            dict_template[prs.slides.index(slide)]['info'][aux[0]]['Multiples'] = True;
                            dict_template[prs.slides.index(slide)]['info'][aux[0]]['objects'] = {};
                        else:
                            dict_template[prs.slides.index(slide)]['info'][aux[0]]['Multiples'] = False;
                            dict_template[prs.slides.index(slide)]['info'][aux[0]]['objects'] = {};
                            dict_template[prs.slides.index(slide)]['info'][aux[0]]['objects']['0'] = shape;

                    if dict_template[prs.slides.index(slide)]['info'][aux[0]]['Multiples']:
                        dict_template[prs.slides.index(slide)]['info'][aux[0]]['objects'][aux[-1]] = shape;



####for i in dict_template:
####    print("{}".format(i))
####    for j in dict_template[i]['info']:
####        print("\t{}".format(j))
######        if dict_template[i][j]['Multiples']:
####        print("{}".format(dict_template[i]['info'][j]))


list_shapeProp = []
for i in dict_template:
    for j in dict_template[i]['info']:
        print(j)
        if dict_template[i]['info'][j]['Multiples']:
            MaxArea = 0;
            for shape_num in dict_template[i]['info'][j]['objects']:
                print(shape_num)
                list_shapeProp.append([])
                x, y, h, w, dict_template[i]['info'][j]['objects'][shape_num] = putData(dict_template[i]['object'], dict_template[i]['info'][j]['objects'][shape_num], j, data);

                list_shapeProp[-1].append(h);
                list_shapeProp[-1].append(w);
                list_shapeProp[-1].append(shape_num);

            for shapeProp in list_shapeProp:
                area = shapeProp[0]*shapeProp[1];
                if MaxArea < area:
                    MaxArea = area;
                    i_MaxArea = shapeProp[2];

            for shapeProp in list_shapeProp:
                if not shapeProp[2] == i_MaxArea:
                    print('Deleting <{}-{}>'.format(j, shapeProp[2]))
                    deleteImage(dict_template[i]['info'][j]['objects'][shapeProp[2]])

            
        else:
            for shape_num in dict_template[i]['info'][j]['objects']:
                x, y, h, w, dict_template[i]['info'][j]['objects'][shape_num] = putData(dict_template[i]['object'], dict_template[i]['info'][j]['objects'][shape_num], j, data);

prs.save("testing1.pptx")
