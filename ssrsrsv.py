#
# ssrsrsv - SQL Server Reporting Services report style validator
#           just because world needs more acronyms
#
# Licence LGPL v2.1
# paavels@gmail.com 2017
#

#!/usr/bin/env python

# TODO:
# textbox widths
# tablix column widths
# tablix height/width
# embeded image

import sys
import xml.etree.ElementTree as ET

default_namespace = 'http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition'
old_namespace = 'http://schemas.microsoft.com/sqlserver/reporting/2008/01/reportdefinition'

ns = {
    'r': default_namespace,
    'rd': 'http://schemas.microsoft.com/SQLServer/reporting/reportdesigner'
    }

def print_help():
    print("This script performs check of SSRS by some arbitrary set of rules.")
    print("")
    print("SSRSRSV.py [pagesize] [filename] [command parameters]")
    print("")
    print("pagesize\t\tSpecifies page size to verify against [a4, a4_landscape, a3_landscape]")
    print("filename\tSpecifies .rdl report definition file to be processed")
    print("")
    print("Commands:")
    print("\tverify\tVerifies report definition against set of rules")

def expected_page_height(pagesize):
    if pagesize == 'a3_landscape':
        return 29.7
    if pagesize == 'a4_landscape':
        return 21
    return 29.7

def expected_page_width(pagesize):
    if pagesize == 'a3_landscape':
        return 42
    if pagesize == 'a4_landscape':
        return 29.7
    return 21

def expected_interactive_height(pagesize):
    return expected_page_height(pagesize) + expected_top_margin(pagesize) + expected_bottom_margin(pagesize)

def expected_interactive_width(pagesize):
    return expected_page_width(pagesize) + expected_left_margin(pagesize) + expected_right_margin(pagesize)

def expected_left_margin(pagesize):
    if pagesize == 'a4':
        return 1.5
    return 0.5

def expected_right_margin(pagesize):
    return 0.5

def expected_top_margin(pagesize):
    if pagesize == 'a4':
        return 0.5
    return 1.5

def expected_bottom_margin(pagesize):
    return 0.5

def expected_body_width(pagesize):
    return expected_page_width(pagesize) - expected_left_margin(pagesize) - expected_right_margin(pagesize)

def expected_header_height():
    return 2

def expected_footer_height():
    return 0.7

def expected_font_face():
    return 'Verdana'

def expected_font_size(documentpart):
    if documentpart == 'header':
        return '7pt'
    if documentpart == 'footer':
        return '7pt'
    if documentpart == 'title':
        return '10pt'

    return '8pt'

def read_report(filename):

    ET.register_namespace('', 'http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition')
    ET.register_namespace('rd', 'http://schemas.microsoft.com/SQLServer/reporting/reportdesigner')

    print("Reading report file:", filename)
    return ET.parse(filename)

def save_report(xml, filename):
    xml.write(filename, 'utf-8', True)

def process_textbox(textbox, documentpart, fix):
    print('  Processing textbox', textbox.attrib['Name'])
    paragraphs = textbox.find('r:Paragraphs', ns).findall('r:Paragraph', ns)

    for paragraph in paragraphs:
        run = paragraph.find('r:TextRuns', ns).find('r:TextRun', ns)

        style = run.find('r:Style', ns)

        font_face = style.find('r:FontFamily', ns)
        font_size = style.find('r:FontSize', ns)

        expected_font = expected_font_face()
        expected_size = expected_font_size(documentpart)

        val = run.find('r:Value', ns)
        if val is not None and val.text and 'ReportName.Value' in val.text:
                expected_size = expected_font_size('title')

        if font_face is not None:
            verify_and_fix_value(font_face, expected_font, 'font face', fix)
        else:
            if not font_face and fix:
                print('   Font face not found, setting font face to', expected_font)
                font_face = ET.SubElement(style, 'FontFamily')
                font_face.text = expected_font

        if font_size is not None:
            verify_and_fix_value(font_size, expected_size, 'font face', fix)
        else:
            if not font_size and fix:
                print('   Font size not found, setting font size to', expected_size)
                font_size = ET.SubElement(style, 'FontSize')
                font_size.text = expected_size                

def process_tablix(tablix, fix):
    print('  Processing tablix', tablix.attrib['Name'])

    rows = tablix.find('r:TablixBody', ns).find('r:TablixRows', ns)
    for row in rows:
        cells = row.find('r:TablixCells', ns)
        for cell in cells:
            items = cell.find('r:CellContents', ns)
            for item in items:
                if 'Textbox' in item.tag:
                    process_textbox(item, 'body', fix)            

def process_report_body(body, fix):
    print('Processing report body')
    items = body.find('r:ReportItems', ns)

    for item in items:
        if 'Textbox' in item.tag:
            process_textbox(item, 'body', fix)
        if 'Tablix' in item.tag:
            process_tablix(item, fix)

def check_old_format(root):
    xmlstr = str(ET.tostring(root))
    if old_namespace in xmlstr:
        print('Detected old namespace. Please upgrade RDL file to 2016 version')
        return False

    return True

def process_body_width(body_width, pagesize, fix):
    print('Processing report body width')
    expected = "{0}cm".format(expected_body_width(pagesize))
    if(body_width.text != expected):
        print("   Expected body width {0}, got {1}".format(expected, body_width.text))
        if(fix):
            print(" + Setting body width to", expected)
            body_width.text = expected

def process_page_footer(page_footer, fix):
    print('Processing report footer')
    footer_height = page_footer.find('r:Height', ns)

    verify_and_fix_value(footer_height, "{0}cm".format(expected_footer_height()), 'footer height', fix)

    items = page_footer.find('r:ReportItems', ns)
    for item in items:
        if 'Textbox' in item.tag:
            process_textbox(item, 'footer', fix)

def process_page_header(page_header, fix):
    print('Processing report header')
    header_height = page_header.find('r:Height', ns)
    expected_height = "{0}cm".format(expected_header_height())

    verify_and_fix_value(header_height, "{0}cm".format(expected_header_height()), 'header height', fix)

    items = page_header.find('r:ReportItems', ns)
    for item in items:
        if 'Textbox' in item.tag:
            process_textbox(item, 'header', fix)

def verify_and_fix_value(elm, expected, label, fix):
    if elm.text != expected:
        print("   Expected {0} of {1}, got {1}".format(label, expected, elm.text))
        if fix:
            print(" + Setting {0} to {1}".format(label, expected))
            elm.text = expected    

def process_page_size(page, pagesize, fix):
    print('Processing page size')

    page_height = page.find('r:PageHeight', ns)
    page_width = page.find('r:PageWidth', ns)
    interactive_height = page.find('r:InteractiveHeight', ns)
    interactive_width = page.find('r:InteractiveWidth', ns)
    left_margin = page.find('r:LeftMargin', ns)
    right_margin = page.find('r:RightMargin', ns)
    top_margin = page.find('r:TopMargin', ns)
    bottom_margin = page.find('r:BottomMargin', ns)

    verify_and_fix_value(page_height, "{0}cm".format(expected_page_height(pagesize)), 'page height', fix)
    verify_and_fix_value(page_width, "{0}cm".format(expected_page_width(pagesize)), 'page width', fix)
    verify_and_fix_value(left_margin, "{0}cm".format(expected_left_margin(pagesize)), 'left margin', fix)
    verify_and_fix_value(right_margin, "{0}cm".format(expected_right_margin(pagesize)), 'right margin', fix)
    verify_and_fix_value(top_margin, "{0}cm".format(expected_top_margin(pagesize)), 'top margin', fix)
    verify_and_fix_value(bottom_margin, "{0}cm".format(expected_bottom_margin(pagesize)), 'bottom margin', fix)

    expected = "{0}cm".format(expected_interactive_height(pagesize))
    if interactive_height is not None:
        verify_and_fix_value(interactive_height, expected, 'interactive height', fix)
    else:
        if not interactive_height and fix:
            print('   Interactive height not found, setting to', expected)
            font_face = ET.SubElement(page, 'InteractiveHeight')
            font_face.text = expected

    expected = "{0}cm".format(expected_interactive_width(pagesize))
    if interactive_width is not None:
        verify_and_fix_value(interactive_width, expected, 'interactive width', fix)
    else:
        if not interactive_width and fix:
            print('   Interactive width not found, setting to', expected)
            font_face = ET.SubElement(page, 'InteractiveWidth')
            font_face.text = expected

def get_body_section(root):
    return root.find('r:ReportSections', ns).find('r:ReportSection', ns)

def get_body_width(body_section):
    return body_section.find('r:Width', ns)

def get_page(body_section):
    return body_section.find('r:Page', ns)

def get_page_header(body_section):
    return get_page(body_section).find('r:PageHeader', ns)

def get_body(body_section):
    return body_section.find('r:Body', ns)

def get_footer(body_section):
    return get_page(body_section).find('r:PageFooter', ns)

def process_report(filename, out_filename, pagesize):
    xml = read_report(filename)
    root = xml.getroot()

    fix_report = True if out_filename else False

    if not check_old_format(root):
        return 1 
    
    body_section = get_body_section(root)

    process_body_width(get_body_width(body_section), pagesize, fix_report)
    process_page_header(get_page_header(body_section), fix_report)
    process_report_body(get_body(body_section), fix_report)
    process_page_size(get_page(body_section), pagesize, fix_report)
    process_page_footer(get_footer(body_section), fix_report)

    if out_filename:
        save_report(xml, out_filename)
        
    return 0

def main(argv):

    pagesize = args[1] if len(argv) > 1 else ''
    filename = args[2] if len(argv) > 2 else ''
    out_filename = args[3] if len(argv) > 3 else ''

    if pagesize and filename:
        return process_report(filename, out_filename, pagesize)

    print_help()

if __name__ == "__main__":

    #args = ['app','a4','test.rdl','test.rdx']
    #main(args)
    main(sys.argv)

