# NPI-Gen
New Product Introduction Presentation Generator

## Objective
Script to automatically generate NPI Submissions given the correct information

## Language
Python

###### Version
```
3.8.10
```

###### Dependencies
```
pip install weasyprint
pip install python-pptx
pip install openpyxl
pip install svglib
pip install reportlab
pip install pywin32
```

## Inputs
> Files which are needed for the script to run
```
Excel file                (NPI_TEMPLATE_FILL_Test.xlsx)
PowerPoint Template       (following the name _Template-xx.pptx_)
Customer Email Template   (CustomerEmail_Template.msg)*
Submission Email Template (SubmissionEmail_Template.msg)*
```

## Outputs
> Files created by the script if succesful execution
```
PowerPoint file   (following the format _NPI-Supplier-PartNumber.pptx_)
Customer Email    (following the format _NPI-Supplier-PartNumber.msg_)
Submission Email  (following the format _Submission-NPI-Supplier-PartNumber.msg_)
```
