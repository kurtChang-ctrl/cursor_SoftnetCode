#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""SQL parameter replacement script for RUNTimeServer.cs"""

import re

# Read the file
file_path = r'c:\04_SoftNet\SoftNet\SoftNetWebII\Services\RUNTimeServer.cs'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

original_len = len(content)

# Replacement patterns for ErrorType 15 with various Time_N_Type values
# Pattern: if (db.DB_SetData($"INSERT... ErrorType='15'...Time_N_Type='X'...
replacements = [
    # Time_N_Type = '4'
    (
        r"if \(db\.DB_SetData\(\$\"INSERT INTO SoftNetSYSDB\.\[dbo\]\.\[APS_Simulation_ErrorData\] \(Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_Type,StationNO\) VALUES \('{[^}]+}','{[^}]+}','{[^}]+}','15','{[^}]+}','{[^}]+}','4','{[^}]+}'\)\"\)\)",
        lambda m: f"if (InsertSimulationError(db, \"15\", simulationId, needId, stationNo: stationNo, timeNType: \"4\", logDate: logDate))"
    ),
]

# Pattern for ErrorType 15 APS_Simulation_ErrorData
pattern_15 = r"if \(db\.DB_SetData\(\$\"INSERT INTO SoftNetSYSDB\.\[dbo\]\.\[APS_Simulation_ErrorData\] \(Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_Type,StationNO\) VALUES \('[^']*',('[^']*'),('[^']*'),('15'),('[^']*'),('[^']*'),('[1234]'),('[^']*')\)\"\)\)"

# Find and replace all ErrorType 15 patterns
pattern_compiled = re.compile(r"if \(db\.DB_SetData\(\$\"INSERT INTO SoftNetSYSDB\.\[dbo\]\.\[APS_Simulation_ErrorData\]", re.MULTILINE)
matches = list(pattern_compiled.finditer(content))
print(f"Found {len(matches)} matches for ErrorType 15 APS_Simulation_ErrorData")

# Manual replacements approach - using exact line replacements
# Line 5903
line_5903_old = """if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_Type,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d2["SimulationId"].ToString()}','15','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','4','{d["StationNO"].ToString()}')"))"""
line_5903_new = """if (InsertSimulationError(db, "15", d2["SimulationId"].ToString(), needId, stationNo: d["StationNO"].ToString(), timeNType: "4", logDate: tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")))"""

content = content.replace(line_5903_old, line_5903_new)
print("Replaced line 5903")

# Continue with other lines...
print(f"Final content length: {len(content)} (original: {original_len})")
print(f"Change in size: {len(content) - original_len}")

# Write back
with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("File saved successfully!")
