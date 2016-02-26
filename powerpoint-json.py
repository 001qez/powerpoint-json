#!/usr/bin/env python

"""powerpoint-json.py:
1. Reads 'product_tags.json' file, then create an empty dictionary.
2. 'product_tags.json' file specify registered product_code (e.g. 'LocalWaters'),
   registered shape_name (e.g. 'TABLE_0') and its expected number of columns,
   and the expected upload_path.
3. Checks the shape_name using both shape.Tags.Item("NAME") and shape.Name
   (When PowerPoint file is edited with PowerPoint 2003, the shape.Name changes,
   so the workaround is to use shape.Tags.Item("NAME").
4. Creates an empty dictionary first, so that it at least tries to produce a valid
   (albeit containing empty text fields).
5. Make a png file for shape_name not equal to TABLE, PAIRS or IGNORE in the
6. Works on the first slide only, so that forecasters can have some reference materials in
   the second slide. Exception is PSOTSL and HADRTSL.
7. Remember to run this thing in Console mode
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

##Comment the next line out for windowless operation
##Application.Visible = True

## Also set the 4th argument to False for windowless operation (FileName,
## ReadOnly, Untitled, WithWindow)
Presentation = Application.Presentations.Open(
    targetpptpath, True, True, False)

def cell_text(cell):
    'return the text in a table cell'
    return cell.Shape.TextFrame.TextRange.Text

def create_empty_pairs():
    temp_dict = collections.OrderedDict()
    temp_dict[''] = ''
    return temp_dict

def create_empty_table(number_of_columns):
    title = ''
    table_array = []
    row_array = []
    for i in range(number_of_columns):
        row_array.append('')
    table_array.append(row_array)
    temp_dict = collections.OrderedDict()
    temp_dict['table_title'] = title
    temp_dict['table_data'] = table_array
    return temp_dict

def create_empty_slide_dict():
    slide_dict = collections.OrderedDict()
    for key in products_details[product_code]['json_items']:
        if key[:5] == 'PAIRS':
            slide_dict[key] = create_empty_pairs()
        if key[:5] == 'TABLE':
            slide_dict[key] = create_empty_table(products_details[product_code]['json_items'][key])
    return slide_dict

def pairs_function(shape, shape_name):
    temp_dict = collections.OrderedDict()
    rows_count = shape.Table.Rows.Count
    cols_count = shape.Table.Columns.Count
    if cols_count != products_details[product_code]['json_items'][shape_name]:
        print 'Warning!'
        print shape_name + ' has ' + str(cols_count) + ' column(s). Expected ' + str(products_details[product_code]['json_items'][shape_name]) + ' column(s).'
    for i in range(1, cols_count, 2):
        for j in range(1, rows_count+1):
            k = cell_text(shape.Table.Cell(j,i))
            v = cell_text(shape.Table.Cell(j,i+1))
            if (k != '') and (v != ''):
                temp_dict[k] = v
    return temp_dict

def table_function(shape, shape_name):
    title = cell_text(shape.Table.Rows(1).Cells.Item(1))
    rows_count = shape.Table.Rows.Count
    cols_count = shape.Table.Columns.Count
    if cols_count != products_details[product_code]['json_items'][shape_name]:
        print 'Warning!'
        print shape_name + ' has ' + str(cols_count) + ' column(s). Expected ' + str(products_details[product_code]['json_items'][shape_name]) + ' column(s).'
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
    slide_dict = create_empty_slide_dict()
    png_export_list = []
    index = 1
    for shape in slide.Shapes:
        
        if shape.Tags.Item('NAME') in products_details[product_code]['json_items']:
            shape_name = shape.Tags.Item('NAME')
        elif shape.Name in products_details[product_code]['json_items']:
            shape_name = shape.Name
        elif shape.Tags.Item('NAME')[:6] == 'IGNORE' or shape.Name[:6] == 'IGNORE':
            shape_name = 'IGNORE'
        else:
            shape_name = shape.Name
            
        if shape_name in products_details[product_code]['json_items']:
            if shape_name[:5] == 'PAIRS':
                slide_dict[shape_name] = pairs_function(shape, shape_name)
            if shape_name[:5] == 'TABLE':
                slide_dict[shape_name] = table_function(shape, shape_name)
        elif shape_name[:6] != 'IGNORE':
            png_export_list.append(index)
        index += 1
    
    if len(png_export_list) > 0:
        slide.Shapes.Range(png_export_list).Export( os.path.join(outputpath, outputfilename+'-'+str(len(ppt_array))+'_'+datetimestamp+'.png'), ppShapeFormatPNG)
    return slide_dict

def slide_function_PSOTSL(slide):
    slide_dict = create_empty_slide_dict()
    png_export_list = []
    index = 1
    for shape in slide.Shapes:
        
        if shape.Tags.Item('NAME') in products_details[product_code]['json_items']:
            shape_name = shape.Tags.Item('NAME')
        elif shape.Name in products_details[product_code]['json_items']:
            shape_name = shape.Name
        elif shape.Tags.Item('NAME')[:6] == 'IGNORE' or shape.Name[:6] == 'IGNORE':
            shape_name = 'IGNORE'
        else:
            shape_name = shape.Name
            
        if shape_name in products_details[product_code]['json_items']:
            if shape_name[:5] == 'PAIRS':
                slide_dict[shape_name] = pairs_function(shape, shape_name)
            if shape_name[:5] == 'TABLE':
                slide_dict[shape_name] = table_function(shape, shape_name)
        elif shape_name[:6] != 'IGNORE':
            png_export_list.append(index)
        index += 1

    if len(png_export_list) > 0:
        slide.Shapes.Range(png_export_list).Export( os.path.join(outputpath, outputfilename+'-'+'0'+'_'+datetimestamp+'.png'), ppShapeFormatPNG)
    return slide_dict

datetimestamp = datetime.now().strftime("%Y%m%d_%H%M")

try:
    with open('products_details.json') as g:
        products_details = json.load(g, object_pairs_hook=collections.OrderedDict)
        g.close()
except:
    print 'There is an error with "products_details.json" file in this executable folder.'
    print os.path.dirname(sys.executable) + ' or in E:\\powerpoint-json\exe '
    print 'This file might be missing or changed incorrectly or corrupted.'
    print 'Restoring this file from the zip archive might fix this error.'
    print ''
    raw_input("Press Enter to close")
    sys.exit()

product_code = outputfilename.split('_')[0]
if not product_code in products_details:
    print 'ERROR!'
    print product_code + ' is not a recognised product name!'
    print 'Recognised product name are:'
    print list(products_details)
    print 'Please check the input PowerPoint filename follows the convention.'
    print ''
    print 'For OverseasSail, OverseasSailWindTemp, etc., '
    print 'the PowerPoint filename should contains only 2 underscore characters.'
    print 'Please check that you do not use space or underscore characters for AreaName/Key'
    print ''
    raw_input("Press Enter to close")
    sys.exit()

with io.open( os.path.join(outputpath, outputfilename+'_'+datetimestamp+'.json') , 'w', encoding='utf8') as f:

    if product_code == 'PSOTSL':
        
        ppt_array = collections.OrderedDict()

        slide = Presentation.Slides(1)
        product_code = 'PSO'
        ppt_array['PSO'] = slide_function_PSOTSL(slide)

        slide = Presentation.Slides(2)
        product_code = 'TSL'
        ppt_array['TSL'] = slide_function_PSOTSL(slide)

        product_code = 'PSOTSL'

    elif product_code == 'HADRTSL':

        ppt_array = collections.OrderedDict()

        slide = Presentation.Slides(1)
        product_code = 'HADR'
        ppt_array['HADR'] = slide_function(slide)

        slide = Presentation.Slides(2)
        product_code = 'TSL'
        ppt_array['TSL'] = slide_function(slide)

        product_code = 'HADRTSL'

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
        slide = Presentation.Slides(1)
        slide_dict = slide_function(slide)
        ppt_array.append(slide_dict)

    ff = json.dumps(ppt_array, indent=4, sort_keys=False, ensure_ascii=False)
    f.write(unicode(ff))
    f.close()

    print 'Please copy and paste files into '
    print products_details[product_code]['upload_path']
    print ''

    # To export the whole ppt to png
    # slide1.Export( <filename here> , "GIF", 1560, 1080)

Application.Quit()
raw_input("Press Enter to close")
