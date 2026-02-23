# BIM–NLP Scheduling Framework

This repository contains the sanitized Python scripts developed for automating 
activity generation, duration estimation, and dependency extraction using a hybrid 
BIM–NLP workflow.

## 📌 What This Code Does
- Extracts activities from BIM-based data
- Generates activity IDs and structured lists
- Estimates durations using productivity rates
- Creates FS/SS/FF/SF logic using rules + SBERT similarity
- Applies heuristic crashing to reduce schedule duration
- Produces Excel sheets ready for Primavera P6 import

## 📂 Files Included
- Activity_ID.py  
- Activity_List.py  
- Activity_Duration.py  
- Generate_Relationships.py  
- Crashing_Duration.py  
- BOQ_Format.py
- BEXPORT ALL ELEMENT FINAL.dyn  
- Pricing.py  
- README.md  

## 🛠 Software Used
- Autodesk Revit  
- Dynamo for Revit  
- Visual Studio Code (Python)  
- Microsoft Excel  
- Primavera P6  

## 🔒 Data Availability
Project BIM data and P6 files are confidential.  
However, sanitized example inputs + code structure are provided for reproducibility.
