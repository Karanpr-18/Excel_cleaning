import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re
import os
import logging

class KadamPlusValidator:
    def __init__(self):
        self.validation_rules = {
            "Student's First Name": ["not null", "no_special_chars"],
            "Student's Age": ["not null", "numeric"],
            "Student's Date of Birth": ["not null", "date"],
            "Father's Name": ["not null", "no_special_chars"],
            "Father's Age": ["not null", "numeric"],
            "Mother's Name": ["not null"],
            "Mother's Age": ["not null", "numeric"],
            "Contact No.": ["not null", "numeric"],
            "House Address": ["not null"],
            "Pincode": ["not null", "numeric"],
            "People living in house": ["not null", "numeric"],
            "Cast": ["not null", "no_special_chars"],
            "Religion": ["not null", "no_special_chars"],
            "Baseline Math": ["not null", "numeric"],
            "Baseline English": ["not null", "numeric"],
            "Baseline EVS": ["not null", "numeric"],
            "Baseline Hindi": ["not null", "numeric"],
            "Baseline Total": ["numeric"],
            "Endline Math": ["numeric"],
            "Endline English": ["numeric"],
            "Endline EVS": ["numeric"],
            "Endline Hindi": ["numeric"],
            "Endline Total": ["numeric"],
            "Grade Test 1": ["numeric"],
            "Grade Test 2": ["numeric"],
            "Grade Test 3": ["numeric"],
            "Grade Test 4": ["numeric"],
            "Grade Test 5": ["numeric"],
            "Enrolment Grade": ["not null"]
        }
        self.red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    def validate_cell(self, value, rules, column_name=None, row=None):
        grade_test_columns = [
            "Grade Test 1", "Grade Test 2", "Grade Test 3", "Grade Test 4", "Grade Test 5"
        ]

        for rule in rules:
            if rule == "not null" and (pd.isnull(value) or str(value).strip() == ""):
                return True, "Value is null or empty"
            if rule == "numeric":
                try:
                    float(value)
                except:
                    return True, "Value is not numeric"
            if rule == "date" and pd.to_datetime(value, errors='coerce') is pd.NaT:
                return True, "Invalid date"
            if rule == "no_special_chars" and not re.match(r"^[A-Za-z ]*$", str(value)):
                return True, "Special characters found"

        if column_name == "Contact No.":
            if pd.notnull(value):
                try:
                    value_str = re.sub(r"\D", "", str(value))  # remove non-digit characters
                    if len(value_str) != 10:
                        return True, "Contact No. must be exactly 10 digits"
                except:
                    return True, "Invalid contact number"

        if column_name in ["Father's Age", "Mother's Age"]:
            try:
                if float(value) < 20:
                    return True, f"{column_name} should not be less than 20"
            except:
                return True, f"Invalid age in {column_name}"

        if column_name == "People living in house":
            try:
                if float(value) <= 1:
                    return True, "People living in house must be more than 1"
            except:
                return True, "Invalid number for People living in house"

        if column_name in grade_test_columns:
            if pd.notnull(value):
                try:
                    score = float(value)
                    if score < 40:
                        return True, "Grade test marks cannot be less than 40"
                except:
                    return True, "Invalid grade test score"

        if column_name in ["Baseline Math", "Baseline English", "Baseline EVS", "Baseline Hindi", "Endline Math", "Endline English", "Endline EVS", "Endline Hindi"]:
            try:
                if column_name.startswith("Baseline") and pd.isnull(value):
                    return True, f"{column_name} cannot be null"
                if pd.notnull(value):
                    if 'Enrolment Grade' in row:
                        grade = str(row['Enrolment Grade']).strip()
                        subject_score = float(value)
                        grade_max = {"1": 10, "2": 20, "3": 30, "4": 40}
                        max_score = grade_max.get(grade)
                        if max_score and subject_score > max_score:
                            return True, f"{column_name} exceeds max allowed ({max_score}) for Grade {grade}"
            except:
                return True, "Invalid subject score"

        if column_name in ["Baseline Total", "Endline Total"]:
            try:
                if column_name == "Baseline Total" and pd.isnull(value):
                    return False, ""
                if pd.notnull(value) and 'Enrolment Grade' in row:
                    grade = str(row['Enrolment Grade']).strip()
                    total_score = float(value)
                    grade_max_total = {"1": 40, "2": 80, "3": 120, "4": 160}
                    max_total = grade_max_total.get(grade)
                    if max_total and total_score > max_total:
                        return True, f"{column_name} exceeds max allowed total ({max_total}) for Grade {grade}"
            except:
                return True, "Invalid total score"

        return False, ""

    def validate_excel(self, file_path, unique_id, download_folder='downloads'):
        try:
            df = pd.read_excel(file_path)
            wb = load_workbook(file_path)
            ws = wb.active

            col_letter_map = {cell.value: cell.column_letter for cell in ws[1] if cell.value in self.validation_rules}
            errors = []

            for idx, row in df.iterrows():
                excel_row = idx + 2
                for col, rules in self.validation_rules.items():
                    if col not in df.columns or col not in col_letter_map:
                        continue
                    value = row[col]
                    has_error, reason = self.validate_cell(value, rules, column_name=col, row=row)
                    if has_error:
                        cell_ref = f"{col_letter_map[col]}{excel_row}"
                        ws[cell_ref].fill = self.red_fill
                        errors.append({"Row": excel_row, "Column": col, "Cell": cell_ref, "Value": value, "Error": reason})

            os.makedirs(download_folder, exist_ok=True)
            validated_path = os.path.join(download_folder, f"{unique_id}_Validated_Output_KadamPlus.xlsx")
            report_path = os.path.join(download_folder, f"{unique_id}_Validation_Report_KadamPlus.xlsx")

            wb.save(validated_path)
            wb.close()

            pd.DataFrame(errors).to_excel(report_path, index=False) if errors else pd.DataFrame(columns=["Row", "Column", "Cell", "Value", "Error"]).to_excel(report_path, index=False)

            logging.info(f"Validation complete. {len(errors)} error(s) found.")
            return {'validated_output': validated_path, 'validation_report': report_path}

        except Exception as e:
            logging.error(f"Validation failed: {str(e)}")
            return None
