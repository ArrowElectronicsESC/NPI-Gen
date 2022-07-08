import win32com.client
import openpyxl
import os

dict_email = {};
Separation_Char = None;

def getExcelData(workbook, sheetName):
    global Separation_Char
    global EmptySpace_Char
    pxl_doc = openpyxl.load_workbook(workbook, data_only = True)
    sheet = pxl_doc[sheetName]
    excelData = {}
    # iterate over columns in the first row
    for i in range(1, sheet.max_column+1):
        excelData[sheet.cell(row=1, column=i).value] = sheet.cell(row=2, column=i).value;

    ## Used for sinchronizaton over variables within the excel file
    sheet = pxl_doc["CONST"];
    Separation_Char = sheet.cell(row = 1, column = 2).value;
    EmptySpace_Char = sheet.cell(row = 2, column = 2).value;

    return excelData

def drawHTMLTable(col_1, col_2, col_3):
    global Separation_Char
    col_1 = col_1.split(Separation_Char);
    col_2 = col_2.split(Separation_Char);
    col_3 = col_3.split(Separation_Char);

    HTML_str = "<table>";
    HTML_str = HTML_str + "\n<tr>";
    HTML_str = HTML_str + "\n<th style = \"width: 30%\">{}</th>\n<th>{}</th>".format(col_1[0], col_2[0], col_3[0]);
    HTML_str = HTML_str + "\n</tr>"
    for i in range(1, 3):
        HTML_str = HTML_str + "\n<tr>";
        HTML_str = HTML_str + "\n<td>{}</td>\n<td>{}</td>".format(col_1[i], col_2[i], col_3[i]);
        HTML_str = HTML_str + "\n</tr>"

    HTML_str = HTML_str + "\n</table>";

##    print(HTML_str);
    return HTML_str;

def func_genSubmissionEmail(data, pptx_name):
##    pptx_name = "NPI-Microchip-DSC612-613.pptx";

    data = getExcelData("./NPI_TEMPLATE_FILL_Test.xlsx", 'TEMPLATE_1_FILL')

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    msg = outlook.OpenSharedItem(r"C:\Users\146294\OneDrive - Arrow Electronics, Inc\Documents\Arrow\Internal Jobs\NPI_Gen\NPI Generator Tool v15\SubmissionEmail_Template.msg")

    msg.Subject = "{} {} - {}".format(msg.Subject, data['Supplier'], data["Title"]);
    msg.Body = msg.Body.replace("Supplier:", "Supplier: " + data['Supplier'])
    msg.Body = msg.Body.replace("Topic:", "Topic: " + data['Title'])
    msg.Body = msg.Body.replace("Application:", "Application:\n" + data['ApplicationText'])
    msg.Attachments.Add(r"C:\Users\146294\OneDrive - Arrow Electronics, Inc\Documents\Arrow\Internal Jobs\NPI_Gen\NPI Generator Tool v15" + "\\{}".format(pptx_name));
    msg.Attachments.Add(r"C:\Users\146294\OneDrive - Arrow Electronics, Inc\Documents\Arrow\Internal Jobs\NPI_Gen\NPI Generator Tool v15\Cust_Email.msg");
     
    msg.SaveAs(r"C:\Users\146294\OneDrive - Arrow Electronics, Inc\Documents\Arrow\Internal Jobs\NPI_Gen\NPI Generator Tool v15\NPI_email.msg", 9)

    del msg, outlook



def func_genClientEmail(data, pptx_name):
##    pptx_name = "NPI-Microchip-DSC612-613.pptx";

    data = getExcelData("./NPI_TEMPLATE_FILL_Test.xlsx", 'TEMPLATE_1_FILL')

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    msg = outlook.OpenSharedItem(r"C:\Users\146294\OneDrive - Arrow Electronics, Inc\Documents\Arrow\Internal Jobs\NPI_Gen\NPI Generator Tool v15\CustomerEmail_Template.msg")

    msg.Subject = "{} - {}".format(data['Supplier'], data["Title"]);
    msg.Body = msg.Body.replace("OverviewText", data['OverviewText'])
    msg.Body = msg.Body.replace("AdvantagesText", data['AdvantagesText'])
    msg.Body = msg.Body.replace("ApplicationText", data['ApplicationText'])
##    msg.Body = msg.Body.replace("PageLink1", '{\field{\*\fldinst { HYPERLINK "{0}" }}{\fldrslt {{1}}}}'.format(data['PageLink1'], data['PageLink1Mask']).encode('utf-8'))
##    Link1 = '{\field{\*\fldinst { HYPERLINK' +  '"' + data['PageLink1'] + '"' + '}}{\fldrslt {' + data['PageLink1Mask'] + '}}}';
##    msg.Body = msg.Body.replace("PageLink1", Link1.encode('utf-8'))
    msg.Attachments.Add(r"C:\Users\146294\OneDrive - Arrow Electronics, Inc\Documents\Arrow\Internal Jobs\NPI_Gen\NPI Generator Tool v15" + "\\{}".format(pptx_name));

##    print(msg.SenderName)
##    print(msg.SenderEmailAddress)
##    print(msg.SentOn)
##    print(msg.To)
##    print(msg.CC)
##    print(msg.BCC)
##    print(msg.Subject)
##    print(msg.Body)

    msg.HTMLBody = msg.HTMLBody.replace("<BR><FONT FACE=\"Calibri\">ImageHere </FONT>", "<BR><img src=\"{}\" alt=\"{}\">".format(os.getcwd() + "\\images\\figures\\" + data['MainFigureImage'], data['Supplier']))
    msg.HTMLBody = msg.HTMLBody.replace("<FONT FACE=\"Calibri\">PageLink1 | PageLink2 </FONT>",
                                        "<a href={}>{}</a><FONT FACE=\"Calibri\"> |  </FONT><a href={}>{}</a>".format(data['PageLink1'], data['PageLink1Mask'],
                                                                                                                                data['PageLink2'], data['PageLink2Mask']))

    ##Table_data = data[""
    msg.HTMLBody = msg.HTMLBody.replace("<BR><FONT FACE=\"Calibri\">TableHere </FONT>", "<BR>{}".format(drawHTMLTable(data['OPNTableColumn1'], data['OPNTableColumn2'], data['OPNTableColumn3'])));
##    msg.HTMLBody = msg.HTMLBody.replace("<FONT FACE=\"Calibri\"> </FONT>", "<BR>{}".format(drawHTMLTable(data['OPNTableColumn1'], data['OPNTableColumn2'], data['OPNTableColumn3'])));
    ##image_index = msg.HTMLBody.rfind("</P>");
    ##
    ##msg.HTMLBody = msg.HTMLBody[:image_index+4] + "\n\n<BR><img src=\"{}\" alt=\"{}\">".format(r"C:\Users\146294\DSC613.PNG", data['Supplier']) + msg.HTMLBody[image_index+4:]

##    print("\n\n")
    print(msg.HTMLBody)
     
    msg.SaveAs(r"C:\Users\146294\OneDrive - Arrow Electronics, Inc\Documents\Arrow\Internal Jobs\NPI_Gen\NPI Generator Tool v15\Cust_email.msg", 3)

