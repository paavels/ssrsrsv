#
# ssrsrsv - SQL Server Reporting Services report style validator
#           just because world needs more acronyms
#
# Licence LGPL v2.1
# paavels@gmail.com 2017
#

#!/usr/bin/env python

import sys
import xml.etree.ElementTree as ET

def print_help():
    print("This script performs check of SSRS by some arbitrary set of rules.")
    print("")
    print("SSRSRSV.py [command] [filename] [command parameters]")
    print("")
    print("command\t\tSpecifies action to be performed on report file")
    print("filename\tSpecifies .rdl report definition file to be processed")
    print("")
    print("Commands:")
    print("\tverify\tVerifies report definition against set of rules")


def read_report(filename):
    return ET.parse(filename)

def save_report(xml, filename):
    xml.write(filename)

def verify_report(filename):
    print("Reading report file:", filename)
    xml = read_report(filename)
    root = xml.getroot()

    ns = {
        'r': 'http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition',
        'rd': 'http://schemas.microsoft.com/SQLServer/reporting/reportdesigner'
        }

    sections = root.find('r:ReportSections', ns)
    body = sections.find('r:ReportSection', ns)
    print(body)

        

def main(argv):

    cmd = args[1] if len(argv) > 1 else ''
    filename = args[2] if len(argv) > 2 else ''

    if cmd == 'verify' and filename:
        verify_report(filename)
        return

    print_help()

if __name__ == "__main__":

    main(args)
    #main(sys.argv)
