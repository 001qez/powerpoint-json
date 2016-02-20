#!/usr/bin/env python

"""powerpoint-json.py: Extract data from a ppt file and outputs
a png file if there is a Group shape in each slide
and a json file for the textual data in the ppt file.
"""

__author__      = "LJS"

import win32com.client, os.path, sys, io, json, collections
from datetime import datetime

import MSO, MSPPT
g = globals()
for c in dir(MSO.constants):    g[c] = getattr(MSO.constants, c)
for c in dir(MSPPT.constants):  g[c] = getattr(MSPPT.constants, c)

## Target powerpoint file
targetpptpath = os.path.abspath(sys.argv[1])
## Output directory
outputpath = os.path.dirname(targetpptpath)
## Output filename
outputfilename = os.path.splitext( os.path.basename(targetpptpath) )[0]

Application = win32com.client.Dispatch("PowerPoint.Application")

## Comment the next line out for windowless operation
## Application.Visible = True

## Also set the 4th argument to False for windowless operation (FileName,
## ReadOnly, Untitled, WithWindow)
Presentation = Application.Presentations.Open(
    targetpptpath, True, True, False)

def cell_text(cell):
    'return the text in a table cell'
    return cell.Shape.TextFrame.TextRange.Text

def title_function(shape):
    temp_dict = collections.OrderedDict()
    temp_dict['slide_title'] = shape.TextFrame.TextRange.Text
    return temp_dict

def pairs_function(shape):
    temp_dict = collections.OrderedDict()
    rows_count = shape.Table.Rows.Count
    cols_count = shape.Table.Columns.Count
    
    for i in range(1, cols_count, 2):
        for j in range(1, rows_count+1):
            k = cell_text(shape.Table.Cell(j,i))
            v = cell_text(shape.Table.Cell(j,i+1))
            if (k != '') and (v != ''):
                temp_dict[k] = v
    return temp_dict

def table_function(shape):
    title = cell_text(shape.Table.Rows(1).Cells.Item(1))
    rows_count = shape.Table.Rows.Count
    table_array = []
    for i in range(2, rows_count+1):
        row_array = []
        row = shape.Table.Rows(i)
        for cell in row.Cells:
            if cell.Shape.HasTextFrame:
                s = cell.Shape.TextFrame.TextRange.Text
                row_array.append(s)
            else:
                row_array.append('')
        table_array.append(row_array)
    temp_dict = collections.OrderedDict()
    temp_dict['table_title'] = title
    temp_dict['table_data'] = table_array
    return temp_dict

def slide_function(slide):
    slide_dict = collections.OrderedDict()
    for shape in slide.Shapes:
        shape_name = shape.Name
        if shape_name[:5] == 'PAIRS':
            slide_dict[shape_name] = pairs_function(shape)
        if shape_name[:5] == 'TABLE':
            slide_dict[shape_name] = table_function(shape)
        if shape_name[:6] == 'EXPORT':
            if shape_name[:6] != 'IGNORE':
                shape.Export( os.path.join(outputpath, outputfilename+'-'+str(len(ppt_array))+'_'+datetimestamp+'.png') , ppShapeFormatPNG)
    return slide_dict

def slide_function_PSOTSL(slide):
    slide_dict = collections.OrderedDict()
    for shape in slide.Shapes:
        shape_name = shape.Name
        if shape_name[:5] == 'PAIRS':
            slide_dict[shape_name] = pairs_function(shape)
        if shape_name[:5] == 'TABLE':
            slide_dict[shape_name] = table_function(shape)
        if shape_name[:6] == 'EXPORT':
            if shape_name[:6] != 'IGNORE':
                shape.Export( os.path.join(outputpath, outputfilename+'-'+'0'+'_'+datetimestamp+'.png') , ppShapeFormatPNG)
    return slide_dict
    

datetimestamp = datetime.now().strftime("%Y%m%d_%H%M")

with io.open( os.path.join(outputpath, outputfilename+'_'+datetimestamp+'.json') , 'w', encoding='utf8') as f:

    product_code = outputfilename.split('_')[0]
    
    if product_code == 'PSOTSL':
        
        ppt_array = collections.OrderedDict()

        slide = Presentation.Slides(1)
        ppt_array['PSO'] = slide_function_PSOTSL(slide)

        slide = Presentation.Slides(2)
        ppt_array['TSL'] = slide_function_PSOTSL(slide)

    elif product_code == 'HADRTSL':

        ppt_array = collections.OrderedDict()

        slide = Presentation.Slides(1)
        ppt_array['HADR'] = slide_function(slide)

        slide = Presentation.Slides(2)
        ppt_array['TSL'] = slide_function(slide)

    elif product_code == 'PSO':

        ppt_array = collections.OrderedDict()

        slide = Presentation.Slides(1)
        ppt_array['PSO'] = slide_function(slide)

    elif product_code == 'HADR':

        ppt_array = collections.OrderedDict()

        slide = Presentation.Slides(1)
        ppt_array['HADR'] = slide_function(slide)

    else:

        ppt_array = []
        
        for slide in Presentation.Slides:
            slide_dict = slide_function(slide)
            ppt_array.append(slide_dict)

    
    ff = json.dumps(ppt_array, indent=4, sort_keys=False, ensure_ascii=False)
    f.write(unicode(ff))
    ## Cannot use the code below this line due to storing unicode issue
    #json.dump(ppt_array, f, indent=4, sort_keys=True, ensure_ascii=False)

# To export the whole ppt to png
# slide1.Export( <filename here> , "GIF", 1560, 1080)

Application.Quit()
