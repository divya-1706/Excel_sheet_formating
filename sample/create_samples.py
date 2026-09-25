"""
Creates realistic sample Excel files for all 5 transformation domains in sample/.
"""

import os
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import pandas as pd
import openpyxl
import config


def create_sample_files():
    os.makedirs(config.SAMPLE_DIR, exist_ok=True)

    # 1. OTIS Test Sequence
    otis_path = os.path.join(config.SAMPLE_DIR, "OTIS_Test_Sequence.xlsx")
    wb_otis = openpyxl.Workbook()
    ws = wb_otis.active
    ws.title = "Test Sequence"
    ws.append(["Sl.No", "Test Name", "Point", "Description", "Voltage", "Current", "Device", "Remarks"])
    otis_rows = [
        ["1.1", "PO SERVICE BRAKE TEST", "P_BAT", "SET Voltage 24V", 24.0, 1.5, "PSU.01", "OK"],
        ["", "", "P30V", "READ THE VALUE IN DMM FOR P30V", "", "", "DMM", ""],
        ["1.2", "MAIN CONTACTOR CHECK", "P_BAT", "SET Voltage 18V", 18.0, 2.0, "PSU.01", "OK"],
        ["", "", "P18V", "READ THE VALUE IN DMM FOR P18V", "", "", "DMM", ""],
        ["2.1", "SAFETY CHAIN TEST", "P_BAT", "SET Voltage 48V", 48.0, 1.0, "PSU.02", "OK"],
        ["", "", "P48V", "READ THE VALUE IN DMM FOR P48V", "", "", "DMM", ""],
    ]
    for r in otis_rows:
        ws.append(r)
    wb_otis.save(otis_path)

    # 2. Student Marks
    student_path = os.path.join(config.SAMPLE_DIR, "Student_Marks_Sheet.xlsx")
    df_student = pd.DataFrame([
        {"USN": "101", "Student Name": "Alice Smith", "Department": "CS", "Semester": 6, "Total Marks": 460},
        {"USN": "102", "Student Name": "Bob Jones", "Department": "EC", "Semester": 6, "Total Marks": 390},
        {"USN": "103", "Student Name": "Charlie Brown", "Department": "CS", "Semester": 6, "Total Marks": 485},
        {"USN": "104", "Student Name": "Diana Prince", "Department": "ME", "Semester": 6, "Total Marks": 310},
    ])
    df_student.to_excel(student_path, index=False)

    # 3. Sales Data
    sales_path = os.path.join(config.SAMPLE_DIR, "Sales_Data.xlsx")
    df_sales = pd.DataFrame([
        {"Invoice No": "INV-2026-01", "Item Description": "Industrial Elevator Relay", "Quantity": 10, "Unit Price": 150.0},
        {"Invoice No": "INV-2026-02", "Item Description": "Safety Brake Switch", "Quantity": 5, "Unit Price": 320.0},
        {"Invoice No": "INV-2026-03", "Item Description": "Power Supply Module 24V", "Quantity": 8, "Unit Price": 210.0},
    ])
    df_sales.to_excel(sales_path, index=False)

    # 4. Attendance
    att_path = os.path.join(config.SAMPLE_DIR, "Employee_Attendance.xlsx")
    df_att = pd.DataFrame([
        {"Employee ID": "EMP-001", "Employee Name": "John Doe", "Department": "Testing", "Days Present": 22, "Days Absent": 0},
        {"Employee ID": "EMP-002", "Employee Name": "Jane Miller", "Department": "Quality Control", "Days Present": 20, "Days Absent": 2},
        {"Employee ID": "EMP-003", "Employee Name": "Robert Taylor", "Department": "Assembly", "Days Present": 21, "Days Absent": 1},
    ])
    df_att.to_excel(att_path, index=False)

    # 5. Inventory
    inv_path = os.path.join(config.SAMPLE_DIR, "Inventory_Stock.xlsx")
    df_inv = pd.DataFrame([
        {"Part No": "P-1001", "Part Description": "PSU 24V Controller Board", "Quantity": 45, "Supplier": "Kaynes Tech", "Unit Cost": 85.0},
        {"Part No": "P-1002", "Part Description": "Waveshare IO Expansion Board", "Quantity": 60, "Supplier": "Waveshare Ltd", "Unit Cost": 45.0},
        {"Part No": "P-1003", "Part Description": "DMM Calibration Module", "Quantity": 15, "Supplier": "Fluke Corp", "Unit Cost": 250.0},
    ])
    df_inv.to_excel(inv_path, index=False)

    print("All 5 sample Excel workbooks generated successfully in sample/")


if __name__ == "__main__":
    create_sample_files()
