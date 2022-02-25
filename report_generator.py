"""

Generate UT report

    input: Excel_file.xlsx
    output: UT_report.html

"""

import os
import xlrd
import json
import shutil
import argparse
from bs4 import BeautifulSoup
import xml.etree.ElementTree as ET
import logging

# HTML color codes
color_na = "FFFFFF"
color_green = "CCFFCC"
color_yellow = "FFFFCC"
color_red = "FFCCCC"
color_gray = "DCDCDC"

JOB_DIR = os.path.abspath('.')
Text_report = ''


def get_subprograms(html_file):
    try:
        try:
            with open(html_file, 'r') as report:
                for line in report:
                    split = line.split('<tr')
                    for el in split:
                        if "grand" in el.lower() and "totals" in el.lower():
                            res = el.split('</td>')[1] + '</td>'
                            return res
                    return "<td><strong>n/a</strong></td>"
        except IndexError:
            return
    except IOError:
        print("NOT FOUND: ", html_file.split("\\")[-1])
        res = "<td><strong>n/a</strong></td>"
        return res


def get_complexity(html_file):
    try:
        try:
            with open(html_file, 'r') as report:
                for line in report:
                    split = line.split('<tr')
                    for el in split:
                        if "grand" in el.lower() and "totals" in el.lower():
                            res = el.split('</td>')[2] + '</td>'
                            return res
                    return "<td><strong>n/a</strong></td>"
        except IndexError:
            return
    except IOError:
        print("NOT FOUND: ", html_file.split("\\")[-1])
        res = "<td><strong>n/a</strong></td>"
        return res


def get_function_coverage(html_file):
    try:
        try:
            with open(html_file, 'r') as report:
                for line in report:
                    split = line.split('<tr')
                    for el in split:
                        if "grand" in el.lower() and "totals" in el.lower():
                            res = el.split('</td>')[3] + '</td>'
                            return res
                    return "<td><strong>n/a</strong></td>"
        except IndexError:
            return
    except IOError:
        print("NOT FOUND: ", html_file.split("\\")[-1])
        res = "<td><strong>n/a</strong></td>"
        return res


def get_function_calls(html_file):
    try:
        try:
            with open(html_file, 'r') as report:
                for line in report:
                    split = line.split('<tr')
                    for el in split:
                        if "grand" in el.lower() and "totals" in el.lower():
                            res = el.split('</td>')[4] + '</td>'
                            return res
                    return "<td><strong>n/a</strong></td>"
        except IndexError:
            return
    except IOError:
        print("NOT FOUND: ", html_file.split("\\")[-1])
        res = "<td><strong>n/a</strong></td>"
        return res


def get_statements(html_file):
    try:
        try:
            with open(html_file, 'r') as report:
                for line in report:
                    split = line.split('<tr')
                    for el in split:
                        if "grand" in el.lower() and "totals" in el.lower():
                            res = el.split('</td>')[5] + '</td>'
                            return res
                    return "<td><strong>n/a</strong></td>"
        except IndexError:
            return
    except IOError:
        print("NOT FOUND: ", html_file.split("\\")[-1])
        res = "<td><strong>n/a</strong></td>"
        return res


def get_branches(html_file):
    try:
        try:
            with open(html_file, 'r') as report:
                for line in report:
                    split = line.split('<tr')
                    for el in split:
                        if "grand" in el.lower() and "totals" in el.lower():
                            res = el.split('</td>')[6] + '</td>'
                            return res
                    return "<td><strong>n/a</strong></td>"
        except IndexError:
            return
    except IOError:
        print("NOT FOUND: ", html_file.split("\\")[-1])
        res = "<td><strong>n/a</strong></td>"
        return res


def get_pairs(html_file):
    try:
        try:
            with open(html_file, 'r') as report:
                for line in report:
                    split = line.split('<tr')
                    for el in split:
                        if "grand" in el.lower() and "totals" in el.lower():
                            res = el.split('</td>')[7] + '</td>'
                            return res
                    return "<td><strong>n/a</strong></td>"
        except IndexError:
            return
    except IOError:
        print("NOT FOUND: ", html_file.split("\\")[-1])
        res = "<td><strong>n/a</strong></td>"
        return res


def write_html(name, store_report_in):
    with open("Cache.json") as json_file:
        DICT = json.load(json_file)

    with open(store_report_in + "\\" + name, 'w+') as report:
        '''Begin HTML'''
        with open(JOB_DIR + '\\template.html', 'r') as template:
            template_text = template.read()

        HTML = ''
        FILTER = ''

        '''Sort by parent Unit'''
        # sorted_by_parent = []

        '''WRITE CONTENTS OF HTML TABLE'''
        for el in DICT:
            number = DICT[el]["number"]
            sofware_asset_group = DICT[el]["sofware_asset_group"]
            software_component = DICT[el]["software_component"]
            software_component_test = DICT[el]["software_component_test"]
            severity_criteria = DICT[el]["severity_criteria"]
            subprograms = DICT[el]["subprograms"]
            complexity = DICT[el]["complexity"]
            function_coverage = DICT[el]["function_coverage"]
            function_calls = DICT[el]["function_calls"]
            statements = DICT[el]["statements"]
            branches = DICT[el]["branches"]
            pairs = DICT[el]["pairs"]
            unit = DICT[el]
            unit_name = software_component
            print('Parsing...', unit_name)
            # print('Parent...', el)

            # HTML += '\n\t<!--=====' + number + '=====-->\n'
            HTML += '\n\t<!--=====' + unit_name.upper() + '=====-->\n'
            tr_classes = ' '.join(['p_' + sofware_asset_group, unit_name + '_master', 'result_row']) + ''

            HTML += '\n\t<tr class="' + tr_classes + '">\n'
            HTML += '\t\t<td bgcolor="#CCD8EE"><strong>' + number + '</strong></td>\n'
            # HTML += '\t\t<td bgcolor="#CCD8EE"><strong>' + component_name + '</strong></td>\n'

            if "ERROR" in sofware_asset_group:
                HTML += '\t\t<td bgcolor="#FFCCCC"><strong>' + sofware_asset_group + '</strong></td>\n'
            else:
                HTML += '\t\t<td bgcolor="#CCD8EE"><strong>' + sofware_asset_group + '</strong></td>\n'
            # HTML += '\t\t<td bgcolor="#CCD8EE"><strong>' + component_name + '</strong></td>\n'
            HTML += '\t\t<td bgcolor="#CCD8EE"><a onclick="test($(this))" class="' + unit_name + 'colapse_btn" ' \
                                                                                                 'href="#">-</a></td' \
                                                                                                 '>\n '
            HTML += '\t\t<td bgcolor="#CCD8EE"><strong>' + unit_name + '</strong></td>\n'

            # Add Main values
            for prop in ['software_component_test', 'severity_criteria', 'subprograms', 'complexity',
                         'function_coverage', 'function_calls', 'statements', 'branches', 'pairs']:
                prop_row = '\t\t' + unit[prop] + '\n'
                # print(prop_row)
                HTML += prop_row

            HTML += '\t</tr>\n'

        template_text = template_text.replace('<[CONTENTS]>', HTML)
        template_text = template_text.replace('<[PARENTS_FILTER]>', FILTER)
        template_text = template_text.replace('<[FOOTER_LINES]>', '''
                <td bgcolor="#FB9316"><b>''' + " " + '''</b></td>
                <td bgcolor="#FB9316" align='center'><b>Summary</b></td>
                <td bgcolor="#FB9316">_</td>
                <td bgcolor="#FB9316" align='center'>''' + str(len(DICT)) + '''</td>
                <td bgcolor="#FB9316" align='center'>''' + " " + '''</td>
                <td bgcolor="#FB9316" align='center'>''' + " " + '''</td>
                <td bgcolor="#FB9316" align='center'>''' + " " + '''</td>
                <td bgcolor="#FB9316" align='center'>''' + " " + '''</td>
                <td bgcolor="#FB9316" align='center'>''' + " " + '''</td>
                <td bgcolor="#FB9316" align='center'>''' + " " + '''</td>
                <td bgcolor="#FB9316" align='center'>''' + " " + '''</td>
                <td bgcolor="#FB9316" align='center'>''' + " " + '''</td>
                <td bgcolor="#FB9316" align='center'>''' + " " + '''</td>
                ''')

        report.write(template_text)


def collect_report_paths(workspace):
    """
    Input: workspace - directory of the workspace with the reports
    Output: list of strings - all the absolute paths for all the  files that ends with index.html

    Implementation reason:
        We can loop through the list for matching names instead of walking though the whole workspace for every test
    """
    every_index_html_file = []
    os.chdir(workspace)

    for root, dirs, files in os.walk(JOB_DIR + f"\\{workspace}"):
        for file in files:
            if str(file).lower().endswith('index.html'):
                every_index_html_file.append(root + "\\" + file)

    return every_index_html_file


def create_text_report(element, function_coverage, function_calls, statements, branches):
    global Text_report
    try:
        Text_report += "------------------------------------------------------------------------"
        Text_report += "\n" + str(element[2]) + "\n"
        Text_report += "Function Coverage: " + str(function_coverage).split('(')[1].split(')')[0] + "\n"
        Text_report += "Function Calls: " + str(function_calls).split('(')[1].split(')')[0] + "\n"
        Text_report += "Statement Coverage: " + str(statements).split('(')[1].split(')')[0] + "\n"
        Text_report += "Branch Coverage: " + str(branches).split('(')[1].split(')')[0] + "\n"
        Text_report += "Not full coverage due to a defensive code which could not be covered." + "\n"
        Text_report += "------------------------------------------------------------------------"
    except:
        pass


def from_sheet_name_get_index(wb, token):
    for i in range(50):
        if str(token).lower() in str(wb.sheet_by_index(i).name).lower():
            print(f"INDEX: {i}")
            return i
    return 0


def should_it_have_UT(name_of_component, path_to_release_note):
    all_sw_components = []
    wb = xlrd.open_workbook(path_to_release_note)
    sheet = wb.sheet_by_index(23)

    for row in range(sheet.nrows):
        all_sw_components.append(sheet.row_values(row))

    for element in all_sw_components:
        if str(name_of_component).lower() in str(element[0]).lower():
            # if cell is empty
            if 0 == len(element[12]):
                return "NA"
            else:
                return element[12]


def read_bom(path_to_bom, path_to_release_note, workspace, store_report_in=JOB_DIR):
    bom_file = path_to_bom
    DICT = {}
    wb = xlrd.open_workbook(bom_file)
    global_element_counter = 1
    all_sw_components = []
    absolute_report_paths = []

    sheet = wb.sheet_by_index(from_sheet_name_get_index(wb, "Text"))
    print(f"Processing BOM sheet: {sheet.name}")
    for row in range(sheet.nrows):
        all_sw_components.append(sheet.row_values(row))

    sheet = wb.sheet_by_index(from_sheet_name_get_index(wb, "Text"))
    print(f"Processing BOM sheet: {sheet.name}")
    for row in range(sheet.nrows):
        all_sw_components.append(sheet.row_values(row))

    is_it_skipped_the_first_row = 0

    for element in all_sw_components:

        """get report paths only once"""
        if "Text" in str(element[5]):
            continue

        if 0 == is_it_skipped_the_first_row:
            absolute_report_paths = collect_report_paths(workspace)
            os.chdir(JOB_DIR)

        if is_it_skipped_the_first_row > 1:
            print(f"Working with element: {element[2]}")
            string_global_element_counter = f"{global_element_counter}"

            software_component_test = should_it_have_UT(f"{element[2].lower()}", path_to_release_note)
            software_component_test = f"<td><strong>{software_component_test}</strong></td>"
            try:
                severity_criteria = f"<td><strong>{int(element[5])}</strong></td>"
            except ValueError as e:
                severity_criteria = f"<td><strong>n/a</strong></td>"

            subprogram = "<td><strong>n/a</strong></td>"
            complexity = "<td><strong>n/a</strong></td>"
            function_coverage = "<td><strong>n/a</strong></td>"
            function_calls = "<td><strong>n/a</strong></td>"
            statements = "<td><strong>n/a</strong></td>"
            branches = "<td><strong>n/a</strong></td>"
            pairs = "<td><strong>n/a</strong></td>"

            for report_path in absolute_report_paths:
                if len(str(element[2]).lower()) != 0 and str(element[2]).lower() in str(report_path).split("\\")[-1].lower():
                    subprogram = get_subprograms(report_path)
                    complexity = get_complexity(report_path)
                    function_coverage = get_function_coverage(report_path)
                    function_calls = get_function_calls(report_path)
                    statements = get_statements(report_path)
                    branches = get_branches(report_path)
                    pairs = get_pairs(report_path)

            print(f"subprogram = {subprogram}")
            print(f"complexity = {complexity}")
            print(f"function_coverage = {function_coverage}")
            print(f"function_calls = {function_calls}")
            print(f"statements = {statements}")
            print(f"branches = {branches}")
            print(f"pairs = {pairs}")

            DICT[string_global_element_counter] = {
                "number": f"{global_element_counter}",
                "sofware_asset_group": f"{element[0]}",
                "software_component": f"{element[2]}",
                "software_component_test": f"{software_component_test}",
                "severity_criteria": f"{severity_criteria}",
                "subprograms": subprogram,
                "complexity": complexity,
                "function_coverage": function_coverage,
                "function_calls": function_calls,
                "statements": statements,
                "branches": branches,
                "pairs": pairs,
            }
            print(f"Data written in json")
            print(f"-------------------------------------")
            create_text_report(element, function_coverage, function_calls, statements, branches)
            global_element_counter += 1
        else:
            is_it_skipped_the_first_row += 1

    print(f"Total reports: {global_element_counter - 1}")
    with open(store_report_in + '\\Result_result.json', 'w') as fp:
        json.dump(DICT, fp)


def get_excel_paths(path_to_bom_dir, path_to_release_note_dir):

    bom = 'Not found'
    rn = 'Not found'

    for root, dirs, files in os.walk(f"{path_to_bom_dir}"):
        for file in files:
            print(str(file).lower())
            if str(file).lower().endswith('.xlsx'):
                bom = root + "\\" + file
    for root, dirs, files in os.walk(f"{path_to_release_note_dir}"):
        for file in files:
            if str(file).lower().endswith('.xlsm'):
                rn = root + "\\" + file

    return bom, rn


def main():
    global Text_report

    """Parsing given arguments"""
    parser = argparse.ArgumentParser()
    parser.add_argument("--path_to_bom", default=JOB_DIR + "\\Excel_file.xlsx",
                        help="Where is located the BOM excel file?")
    parser.add_argument("--path_to_release_note", default=JOB_DIR + "\\Excel_file.xlsx",
                        help="Where is located the BOM excel file?")
    parser.add_argument("--workspace", help="Where is the workspace located?")
    parser.add_argument("--store_report_in", default=JOB_DIR, help="Where to store the UT report?")
    args = parser.parse_args()
    path_to_bom_dir = args.path_to_bom
    path_to_release_note_dir = args.path_to_release_note

    path_to_bom, path_to_release_note = get_excel_paths(path_to_bom_dir, path_to_release_note_dir)

    workspace = args.workspace
    store_report_in = args.store_report_in

    '''Read BOM and format html reports -> generate the final UT html report'''
    read_bom(path_to_bom, path_to_release_note, workspace, store_report_in)
    write_html("result.html", store_report_in)
    with open("UT_values.txt", "w") as text_report:
        text_report.write(Text_report)


if __name__ == '__main__':
    main()
