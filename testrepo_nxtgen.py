import os
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

class TestRepo:
    def __init__(self):
        pass

    def parse_xml_file(self, file_path):
        tree = ET.parse(file_path)
        root = tree.getroot()
        testgroups = root.findall('testgroup')
        result_rows = []

        for testgroup in testgroups:
            title = testgroup.find('title').text
            result_rows.append([f"{title}"])

            testcases = testgroup.findall('testcase')
            pass_count = 0
            fail_count = 0

            for testcase in testcases:
                title = testcase.find('title').text

                has_fail = False
                teststeps = testcase.findall('.//teststep')
                for teststep in teststeps:
                    result = teststep.attrib.get('result', 'pass')
                    if result == 'fail':
                        has_fail = True
                        break

                verdict_element = testcase.find('verdict')
                verdict_result = verdict_element.attrib.get('result', 'pass') if verdict_element is not None else 'pass'

                checkstatistic_element = testcase.find('.//checkstatistic')
                checkstatistic_result = checkstatistic_element.attrib.get('result', 'pass') if checkstatistic_element is not None else 'pass'

                failure_ratio = None

                if not has_fail and verdict_result == 'pass' and checkstatistic_result == 'pass':
                    pass_count += 1
                else:
                    if checkstatistic_element is not None:
                        xinfo_elements = checkstatistic_element.findall('.//xinfo')
                        for xinfo_element in xinfo_elements:
                            name_element = xinfo_element.find('name')
                            if name_element is not None and name_element.text == "Failure ratio (in %)":
                                description_element = xinfo_element.find('description')
                                if description_element is not None:
                                    failure_ratio = description_element.text
                    failute_ratiotic = failure_ratio.split("%")[0]                
                    result_rows.append([title, "Fail", failute_ratiotic])
                    fail_count += 1

        return result_rows
    
    def get_summary_counts(self, result_rows):
        # Calculate total, pass, and fail counts
        total_count = len(result_rows) - 4
        pass_count = sum(1 for row in result_rows if len(row) > 0 and row[-1] == "Pass")
        fail_count = sum(1 for row in result_rows if len(row) > 0 and row[-1] == "Fail")
        return total_count, pass_count, fail_count

    def write_summary(self, ws, total_count, pass_count, fail_count):
        # Write the summary to the worksheet
        ws.append(["Total Test Case Titles:", total_count])
        ws.append(["Pass Count:", pass_count])
        ws.append(["Fail Count:", fail_count])
        ws.append([])

    def adjust_column_widths(self, ws):
        # Adjust column widths to fit content
        for column in ws.columns:
            max_length = 0
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            column_letter = get_column_letter(column[0].column)
            ws.column_dimensions[column_letter].width = adjusted_width

    def create_filtered_sheet(self, wb, ws):
        # Create a new filtered sheet and copy non-Pass rows
        ws2 = wb.create_sheet(title="Filtered")
        for row in ws.iter_rows():
            if len(row) > 1 and row[1].value != "Pass":
                ws2.append([cell.value for cell in row])
                for new_cell, old_cell in zip(ws2[ws2.max_row], row):
                    if old_cell.has_style:
                        new_cell.font = old_cell.font.copy()
                        new_cell.border = old_cell.border.copy()
                        new_cell.fill = old_cell.fill.copy()
                        new_cell.number_format = old_cell.number_format
                        new_cell.protection = old_cell.protection.copy()
                        new_cell.alignment = old_cell.alignment.copy()
        return ws2


    # In the TestRepo class, modify the process_xml_files method as follows
    def process_xml_files(self, folder_path):
        output_file = "test_results.xlsx"
        # Check if the output file already exists
        if os.path.exists(output_file):
            # Load the existing workbook
            wb = load_workbook(output_file)
            # Use the existing worksheet
            ws = wb.active
        else:
            # If the output file does not exist, create a new workbook and worksheet
            wb = Workbook()
            ws = wb.active

        red_fill = PatternFill(start_color="ED583A", end_color="ED583A", fill_type="solid")
        title_fill = PatternFill(start_color="B3B3B3", end_color="B3B3B3", fill_type="solid")
        subtitle_fill = PatternFill(start_color="729E84", end_color="729E84", fill_type="solid")
        bold_font = Font(bold=True)
        white_font = Font(color="FFFFFF")

        for filename in os.listdir(folder_path):
            if filename.endswith(".xml"):
                file_path = os.path.join(folder_path, filename)
                group_name = filename.replace("_test_report.xml", "")
                ws.append([f"{group_name}"])
                ws.cell(row=ws.max_row, column=1).font = bold_font
                ws.cell(row=ws.max_row, column=1).fill = title_fill

                result_rows = self.parse_xml_file(file_path)

                for row in result_rows:
                    ws.append(row)
                    if len(row) > 1 and row[1] == "Fail":
                        ws.cell(row=ws.max_row, column=1).fill = red_fill
                        ws.cell(row=ws.max_row, column=1).font = white_font
                    elif len(row) > 0 and row[0].startswith("XH8"):
                        ws.cell(row=ws.max_row, column=1).fill = subtitle_fill
                        ws.cell(row=ws.max_row, column=1).font = bold_font

                        

                total_count, pass_count, fail_count = self.get_summary_counts(result_rows)
                self.write_summary(ws, total_count, pass_count, fail_count)
                ws.append([])

        self.adjust_column_widths(ws)

        # Save the updated workbook to the output file
        wb.save(output_file)



if __name__ == "__main__":
    test_repo = TestRepo()
    folder_path = r"C:\\Users\\ROG\\Desktop\\autorepo"
    
    test_repo.process_xml_files(folder_path)


