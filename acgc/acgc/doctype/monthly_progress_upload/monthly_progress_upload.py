# Copyright (c) 2026, Divya and contributors
# For license information, please see license.txt

# import frappe
from frappe.model.document import Document


class MonthlyProgressUpload(Document):
	pass
    
import re
import pandas as pd

def get_numeric_value(value):
    if pd.isna(value):
        return 0

    if isinstance(value, (int, float)):
        return float(value)

    match = re.search(r"[-+]?\d*\.?\d+", str(value))
    return float(match.group()) if match else 0

import frappe
import pandas as pd
from frappe.utils import getdate
from datetime import datetime

@frappe.whitelist()
def import_monthly_progress(upload_doc):
    upload_doc = frappe.get_doc("Monthly Progress Upload", upload_doc)
    if not upload_doc.upload:
        frappe.throw("Please upload an Excel file.")
    site = upload_doc.select_site
    posting_date = getdate(upload_doc.posting_date)
    month = posting_date.month
    year = posting_date.year
    try:
        file_doc = frappe.get_doc("File",{"file_url": upload_doc.upload})
        file_path = file_doc.get_full_path()
    except Exception:
        frappe.throw("Unable to locate uploaded file.")
    try:
        df = pd.read_excel(file_path,sheet_name="KPI Title")
    except Exception as e:
        frappe.throw(f"Error reading Excel file: {str(e)}")
    existing = frappe.db.sql("""
        SELECT name
        FROM `tabMonthly Progress Data`
        WHERE site=%s
        AND MONTH(posting_date)=%s
        AND YEAR(posting_date)=%s
        LIMIT 1
    """, (site, month, year), as_dict=True)
    if existing:
        doc = frappe.get_doc(
            "Monthly Progress Data",
            existing[0].name
        )
        doc.set("kpi_title", [])
    else:
        doc = frappe.new_doc("Monthly Progress Data")
        doc.site = site
        doc.posting_date = posting_date
    for _, row in df.iterrows():
        kpi_title = row.get("KPI Title")
        performance_measure = row.get("Performance Measure")
        achieved = row.get("Achieved")
        performance_measure = (
            0 if pd.isna(performance_measure)
            else float(performance_measure) * 100
        )

        achieved = (
            0 if pd.isna(achieved)
            else float(achieved) * 100
        )
        if pd.isna(kpi_title):
            continue
        doc.append("kpi_title", {
            "kpi_title": str(kpi_title).strip(),
            "performance_measure": round(performance_measure, 2),
            "achieved": round(achieved, 2)
        })
    # Training On
    try:
        training_df = pd.read_excel(file_path,sheet_name="Training On")
    except Exception:
        training_df = None
    doc.set("training_on", [])
    if training_df is not None:
        training_df.columns = [
            "training_on",
            "planned",
            "conducted",
            "achieved"
        ]
        training_df = training_df.iloc[1:]
        for _, row in training_df.iterrows():

            training_on = row.get("training_on")

            if pd.isna(training_on):
                continue

            planned = 0 if pd.isna(row.get("planned")) else float(row.get("planned"))
            conducted = row.get("conducted")
            if pd.isna(conducted):
                conducted = ""
            elif isinstance(conducted, (pd.Timestamp, datetime)):
                conducted = conducted.strftime("%d-%m-%Y %H:%M:%S")
            else:
                conducted = str(conducted)

            achieved = 0 if pd.isna(row.get("achieved")) else float(row.get("achieved")) * 100

            doc.append("training_on", {
                "training_on": str(training_on).strip(),
                "planned": planned,
                "conducted": conducted,
                "achieved_": round(achieved, 2)
            })
    try:
        safety_df = pd.read_excel(file_path,sheet_name="Safety & Environment Checklist",header=None)
        doc.set("safety_and_environment_checklist", [])
        for _, row in safety_df.iloc[3:].iterrows():
            summary = row[0]
            applicable = row[1]
            actual = row[2]
            satisfactory_rating = row[3]

            if pd.isna(summary):
                continue
            if pd.isna(satisfactory_rating):
                rating = 0
            else:
                rating = round(float(satisfactory_rating) * 100, 2)
            if str(summary).strip()!="OVERALL SCORE & PERCENTAGE":
                doc.append("safety_and_environment_checklist", {
                    "summary_of_scores_and_percentage": str(summary).strip(),
                    "applicable": 0 if pd.isna(applicable) else applicable,
                    "actual": 0 if pd.isna(actual) else actual,
                    "satisfactory_rating": rating
                })
    except Exception as e:
        frappe.log_error(frappe.get_traceback(),f"Safety Sheet Import Error: {str(e)}")
    # Gasoline Consumption
    try:
        gasoline_df = pd.read_excel(file_path,sheet_name="Gasoline Consumption",header=None)
        doc.set("gasoline_consumption", [])
        for _, row in gasoline_df.iloc[1:].iterrows():

            if pd.isna(row[0]):
                continue
            if str(row[0]).strip().upper() == "DATE":
                continue
            try:
                date = pd.to_datetime(row[0]).date()
            except Exception:
                continue

            gasoline_91 = row[1]
            diesel = row[2]
            date = ""
            if not pd.isna(row[0]):
                date = pd.to_datetime(row[0]).strftime("%d-%m-%Y")
            doc.append("gasoline_consumption", {
                "date": date,
                "gasoline_91": 0 if pd.isna(gasoline_91) else gasoline_91,
                "diesel": 0 if pd.isna(diesel) else diesel
            })
    except Exception as e:
        frappe.log_error(frappe.get_traceback(),f"Gasoline Consumption Import Error: {str(e)}")
    # Sewage Water & Garbage
    try:
        sewage_df = pd.read_excel(file_path,sheet_name="Sewage water & Garbage ",header=None,skiprows=2 )
        doc.set("sewage_water_and_garbage", [])
        for _, row in sewage_df.iterrows():
            if pd.isna(row[0]):
                continue
            try:
                date = pd.to_datetime(row[0]).date()
            except Exception:
                continue

            # sewage_water = row[1]
            # garbage_collection = row[2]

            # doc.append("sewage_water_and_garbage", {
            #     "date": date,
            #     "sewage_water": 0 if pd.isna(sewage_water) else float(sewage_water),
            #     "garbage_collection": 0 if pd.isna(garbage_collection) else float(garbage_collection)
            # })
            sewage_water = get_numeric_value(row[1])
            garbage_collection = get_numeric_value(row[2])

            doc.append("sewage_water_and_garbage", {
                "date": date,
                "sewage_water": round(sewage_water,2),
                "garbage_collection": round(garbage_collection,2)
            })

    except Exception as e:
        frappe.log_error(frappe.get_traceback(),f"Sewage Water & Garbage Import Error: {str(e)}")
    # Water Consumption
    try:
        water_df = pd.read_excel(file_path,sheet_name="Water Consumption",skiprows=1 )

        water_df.columns = water_df.columns.str.strip()
        doc.set("water_consumption", [])

        # case-insensitive column lookup (handles "RAW  Water", "Raw Water", "raw water" etc.)
        col_map = {re.sub(r"\s+"," ",str(c)).strip().lower(): c for c in water_df.columns}
        raw_col = col_map.get("raw water")
        sweet_col = col_map.get("sweet water for ps & rc") or col_map.get("sweet water")

        for _, row in water_df.iterrows():
            if pd.isna(row.get("DATE")):
                continue

            try:
                date = pd.to_datetime(row["DATE"]).date()
            except Exception:
                continue 

            raw_water_val = row.get(raw_col, 0) if raw_col else 0
            sweet_water_val = row.get(sweet_col, 0) if sweet_col else 0

            raw_water = 0 if pd.isna(raw_water_val) else round(float(raw_water_val),2)
            sweet_water = 0 if pd.isna(sweet_water_val) else round(float(sweet_water_val),2)

            doc.append("water_consumption", {
                "date": date,
                "raw_water": raw_water,
                "sweet_water": sweet_water
            })
            
    except Exception as e:
        frappe.log_error(frappe.get_traceback(),f"Water Consumption Import Error: {str(e)}")
    # Energy Management
    try:
        water_df = pd.read_excel(file_path,sheet_name="Energy Management",skiprows=2 )
        water_df.columns = water_df.columns.str.strip()
        doc.set("energy_management", [])
        for _, row in water_df.iterrows():
            if pd.isna(row.get("DATE")):
                continue
            try:
                date = pd.to_datetime(row["DATE"]).date()
            except Exception:
                continue 

            unit_consumed_val = row.get("Unit Consumed", 0)

            unit_consumed = 0 if pd.isna(unit_consumed_val) else float(round(unit_consumed_val,0))

            doc.append("energy_management", {
                "date": date,
                "unit_consumed": unit_consumed,
            })
            
    except Exception as e:
        frappe.log_error(frappe.get_traceback(),f"Energy Management Import Error: {str(e)}")
    # Safety
    try:
        water_df = pd.read_excel(file_path,sheet_name="Safety",skiprows=1)
        water_df.columns = water_df.columns.str.strip()
        doc.set("safety", [])
        for _, row in water_df.iterrows():
            injuiry_recorded_val = row.get("Injuiry Recorded", "")
            count_val=row.get("Count","")
            injuiry_recorded = "" if pd.isna(injuiry_recorded_val) else str(injuiry_recorded_val)
            count = 0 if pd.isna(count_val) else float(count_val)
            doc.append("safety", {
                "injuiry_recorded": injuiry_recorded,
                "count":count
            })
            
    except Exception as e:
        frappe.log_error(frappe.get_traceback(),f"Safety Import Error: {str(e)}")
    # Saudisation
    try:
        water_df = pd.read_excel(file_path,sheet_name="SAUDISATION",skiprows=1)
        water_df.columns = water_df.columns.str.strip()
        doc.set("saudisation", [])

        for _, row in water_df.iterrows():
            functions_val = row.get("Functions", "")
            total_val=row.get("Total",0)
            saudi_nationals_val=row.get("Saudi Nationals",0)
            functions = "" if pd.isna(functions_val) else str(functions_val)
            total = 0 if pd.isna(total_val) else float(total_val)
            saudi_nationals= 0 if pd.isna(saudi_nationals_val) else float(saudi_nationals_val)
            if functions!="Total Manpower vs Saudi National":
                doc.append("saudisation", {
                    "functions": functions,
                    "total":total,
                    "saudi_nationals":saudi_nationals
                })
            
    except Exception as e:
        frappe.log_error(frappe.get_traceback(),f"SAUDISATION Import Error: {str(e)}")
    try:
        mp_df = pd.read_excel(file_path,sheet_name="Manpower",header=None)
        doc.set("manpower_summary", [])
        doc.set("manpower_detailed", [])
        summary_idx = None
        detailed_idx = None

        for idx, row in mp_df.iterrows():
            first_col = str(row[0]).strip() if not pd.isna(row[0]) else ""

            if first_col == "Manpower Summary":
                summary_idx = idx

            elif first_col == "Manpower Detailed":
                detailed_idx = idx
        if summary_idx is not None and detailed_idx is not None:

            summary_df = mp_df.iloc[summary_idx + 2 : detailed_idx]

            for _, row in summary_df.iterrows():

                function = row[0]

                if pd.isna(function):
                    continue

                indicative = 0 if pd.isna(row[1]) else float(row[1])
                actual = 0 if pd.isna(row[2]) else float(row[2])

                doc.append("manpower_summary", {
                    "function": str(function).strip(),
                    "indicative": indicative,
                    "actual": actual
                })      
        if detailed_idx is not None:
            detailed_df = mp_df.iloc[detailed_idx + 2 :]
            for _, row in detailed_df.iterrows():
                function = row[0]
                if pd.isna(function):
                    continue
                indicative = 0 if pd.isna(row[1]) else float(row[1])
                actual = 0 if pd.isna(row[2]) else float(row[2])
                if str(function).strip()!="TOTAL":
                    doc.append("manpower_detailed", {
                        "function": str(function).strip(),
                        "indicative": indicative,
                        "actual": actual
                    })

    except Exception as e:
        frappe.log_error(frappe.get_traceback(),f"Manpower Import Error: {str(e)}")
    # Resident Satisfaction Survey
    try:
        rss_df = pd.read_excel(file_path,sheet_name="Resident Satisfaction Survey",header=None)
        doc.set("resident_satisfaction_survey", [])
        header_idx = None
        for idx, row in rss_df.iterrows():
            first_col = str(row[0]).strip() if not pd.isna(row[0]) else ""

            if first_col == "Attribute":
                header_idx = idx
                break

        if header_idx is not None:

            survey_df = rss_df.iloc[header_idx + 1:]

            for _, row in survey_df.iterrows():

                attribute = row[0]

                if pd.isna(attribute):
                    continue

                week_1 = 0 if pd.isna(row[1]) else round(float(row[1]), 2)
                week_2 = 0 if pd.isna(row[2]) else round(float(row[2]), 2)
                week_3 = 0 if pd.isna(row[3]) else round(float(row[3]), 2)
                week_4 = 0 if pd.isna(row[4]) else round(float(row[4]), 2)

                doc.append("resident_satisfaction_survey", {
                    "attribute": str(attribute).strip(),
                    "week_1": week_1,
                    "week_2": week_2,
                    "week_3": week_3,
                    "week_4": week_4
                })

    except Exception as e:
        frappe.log_error(frappe.get_traceback(),f"Resident Satisfaction Survey Import Error: {str(e)}")
    # Customer Satisfaction Survey
    try:
        customer_df = pd.read_excel(
            file_path,
            sheet_name="Customer Satisfaction Survey",
            header=None
        )

        doc.set("customer_satisfaction_survey", [])

        header_idx = None

        for idx, row in customer_df.iterrows():

            values = [
                str(v).strip()
                for v in row
                if not pd.isna(v)
            ]

            if "Resident Name" in values:
                header_idx = idx
                break

        if header_idx is not None:

            headers = customer_df.iloc[header_idx].fillna("").astype(str).str.strip()

            survey_df = customer_df.iloc[header_idx + 1:].copy()
            survey_df.columns = headers

            column_map = {
                "Resident Name": "resident_name",
                "Room Number": "room_number",
                "How do you rate our employees responsiveness with you": "employees_responsiveness",
                "How do you rate front desk employee": "rate_desk_employee",
                "How do you rate room service employees": "room_service_employees",
                "How do you rate our employees professionalism with you": "employees_professionalism",
                "How do you rate service quality": "service_quality",
                "How do you rate the cleanliness of the room": "cleanliness_of_the_room",
                "How do you rate the Technical Service Quality": "technical_service_quality",
                "Taste of the food": "taste_of_the_food",
                "Service": "service",
                "Food quality": "food_quality",
                "Restaurant employees service quality": "restaurant_employees_service_quality",
                "Laundry Service Quality": "laundry_service_quality"
            }

            for _, row in survey_df.iterrows():

                resident_name = row.get("Resident Name")

                if pd.isna(resident_name):
                    continue

                data = {}

                for excel_col, fieldname in column_map.items():

                    value = row.get(excel_col, "")

                    data[fieldname] = (
                        ""
                        if pd.isna(value)
                        else str(value).strip()
                    )

                doc.append(
                    "customer_satisfaction_survey",
                    data
                )

    except Exception as e:
        frappe.log_error(
            frappe.get_traceback(),
            f"Customer Satisfaction Survey Import Error: {str(e)}"
        )
    # try:
    #     customer_df = pd.read_excel(file_path,sheet_name="Customer Satisfaction Survey",header=None)
    #     doc.set("customer_satisfaction_survey", [])
    #     header_idx = None
    #     for idx, row in customer_df.iterrows():
    #         first_col = str(row[0]).strip() if not pd.isna(row[0]) else ""

    #         if first_col == "Resident Name":
    #             header_idx = idx
    #             break

    #     if header_idx is not None:

    #         survey_df = customer_df.iloc[header_idx + 1:]

    #         for _, row in survey_df.iterrows():

    #             resident_name = row[0]

    #             if pd.isna(resident_name):
    #                 continue

    #             doc.append("customer_satisfaction_survey", {
    #                 "resident_name": str(resident_name).strip(),
    #                 "room_number": "" if pd.isna(row[1]) else str(row[1]),
    #                 "employees_responsiveness": "" if pd.isna(row[2]) else str(row[2]).strip(),
    #                 "rate_desk_employee": "" if pd.isna(row[3]) else str(row[3]).strip(),
    #                 "room_service_employees": "" if pd.isna(row[4]) else str(row[4]).strip(),
    #                 "employees_professionalism": "" if pd.isna(row[5]) else str(row[5]).strip(),
    #                 "service_quality": "" if pd.isna(row[6]) else str(row[6]).strip(),
    #                 "cleanliness_of_the_room": "" if pd.isna(row[7]) else str(row[7]).strip(),
    #                 "technical_service_quality": "" if pd.isna(row[8]) else str(row[8]).strip(),
    #                 "taste_of_the_food": "" if pd.isna(row[9]) else str(row[9]).strip(),
    #                 "service": "" if pd.isna(row[10]) else str(row[10]).strip(),
    #                 "food_quality": "" if pd.isna(row[11]) else str(row[11]).strip(),
    #                 "restaurant_employees_service_quality": "" if pd.isna(row[12]) else str(row[12]).strip(),
    #                 "laundry_service_quality": "" if pd.isna(row[13]) else str(row[13]).strip(),
    #             })

    # except Exception as e:
    #     frappe.log_error(frappe.get_traceback(),f"Customer Satisfaction Survey Import Error: {str(e)}")
    # Catering and Laundry Services
    try:
        rss_df = pd.read_excel(
            file_path,
            sheet_name="Catering and Laundry Services",
            header=None
        )
        doc.set("catering_and_laundry_services", [])
        header_idx = None
        for idx, row in rss_df.iterrows():
            first_col = str(row[0]).strip() if not pd.isna(row[0]) else ""

            if first_col == "Attribute":
                header_idx = idx
                break

        if header_idx is not None:

            survey_df = rss_df.iloc[header_idx + 1:]

            for _, row in survey_df.iterrows():

                attribute = row[0]
                if attribute == "Meal":
                    break

                if pd.isna(attribute):
                    continue
                

                week_1 = 0 if pd.isna(row[1]) else float(row[1])
                week_2 = 0 if pd.isna(row[2]) else float(row[2])
                week_3 = 0 if pd.isna(row[3]) else float(row[3])
                week_4 = 0 if pd.isna(row[4]) else float(row[4])

                doc.append("catering_and_laundry_services", {
                    "attribute": str(attribute).strip(),
                    "week_1": week_1,
                    "week_2": week_2,
                    "week_3": week_3,
                    "week_4": week_4
                })
        for _, row in rss_df.iterrows():

            value = str(row[0]).strip() if not pd.isna(row[0]) else ""

            if value == "Breakfast":
                doc.breakfast = 0 if pd.isna(row[1]) else float(row[1])

            elif value == "Lunch":
                doc.lunch = 0 if pd.isna(row[1]) else float(row[1])

            elif value == "Dinner":
                doc.dinner = 0 if pd.isna(row[1]) else float(row[1])

            # elif value == "Laundry":
            #     doc.laundry = 0 if pd.isna(row[1]) else float(row[1])
            elif value == "Laundry":
                doc.laundry = get_numeric_value(row[1])

    except Exception as e:
        frappe.log_error(frappe.get_traceback(),f"Resident Satisfaction Survey Import Error: {str(e)}")
    # Occupancy Report
    try:
        occ_df = pd.read_excel(file_path,sheet_name="Occupancy Report",header=None)

        doc.set("occupancy_report", [])
        doc.set("occupancy_category_report", [])

        building_header = None
        category_header = None

        for idx, row in occ_df.iterrows():

            first_col = str(row[0]).strip() if not pd.isna(row[0]) else ""

            if first_col == "BUILDING NO":
                building_header = idx

            elif first_col == "Category" and category_header is None:
                category_header = idx

        def num(row, i, pct=False):
            if i is None or pd.isna(row[i]):
                return 0
            try:
                v = float(row[i])
            except (ValueError, TypeError):
                return 0
            return v * 100 if pct else v

        header_words = {"BUILDING NO", "CATEGORY", "OCCUPANCY REPORT"}

        def table_end(start):
            for idx in range(start, len(occ_df)):
                c0 = occ_df.iloc[idx, 0]
                c0 = "" if pd.isna(c0) else str(c0).strip().upper()
                if c0 == "TOTAL":
                    return idx + 1
                if c0 in header_words:
                    return idx
            return len(occ_df)

        if building_header is not None:

            headers = (
                occ_df.iloc[building_header]
                .fillna("")
                .astype(str)
                .str.upper()
                .str.strip()
            )

            def bcol(*keys):
                matches = [i for i, h in headers.items() if any(k in h for k in keys)]
                return matches[-1] if matches else None

            total_col = bcol("TOTAL")
            occupied_col = bcol("OCCUPIED")
            vacant_col = bcol("VACANT")
            occupancy_col = bcol("OCCUPANCY")
            type_col = bcol("TYPE OF BUILDING")

            building_df = occ_df.iloc[building_header + 1 : table_end(building_header + 1)]

            for _, row in building_df.iterrows():

                building_number = row[0]

                if pd.isna(building_number):
                    continue

                if str(building_number).strip().upper() == "TOTAL":

                    doc.total_rooms_occupancy_wise = round(num(row, total_col), 2)

                    doc.total_occupied_rooms = round(num(row, occupied_col), 2)

                    doc.total_vacant_rooms = round(num(row, vacant_col), 2)

                    doc.total_occupancy_percentage = round(num(row, occupancy_col, pct=True), 2)

                    continue

                doc.append("occupancy_report", {
                    "building_number": str(building_number).strip(),
                    "total_rooms": num(row, total_col),
                    "occupied_rooms": num(row, occupied_col),
                    "vacant_rooms": num(row, vacant_col),
                    "occupancy_": num(row, occupancy_col, pct=True),
                    "type_of_building": "" if type_col is None or pd.isna(row[type_col]) else str(row[type_col]).strip()
                })
        if category_header is not None:

            cat_headers = (
                occ_df.iloc[category_header]
                .fillna("")
                .astype(str)
                .str.upper()
                .str.strip()
            )

            def ccol(*keys):
                matches = [i for i, h in cat_headers.items() if any(k in h for k in keys)]
                return matches[0] if matches else None

            c_total = ccol("TOTAL")
            c_occupied = ccol("OCCUPIED")
            c_vacant = ccol("VACANT")
            c_occupancy = ccol("OCCUPANCY")

            category_df = occ_df.iloc[category_header + 1 : table_end(category_header + 1)]

            for _, row in category_df.iterrows():

                category = row[0]

                if pd.isna(category):
                    continue

                if str(category).strip().upper() == "TOTAL":

                    doc.total_rooms_category_wise = round(num(row, c_total), 2)

                    doc.total_occupied_rooms_category_wise = round(num(row, c_occupied), 2)

                    doc.total_vacant_rooms_category_wise = round(num(row, c_vacant), 2)

                    doc.total_occupancy_percentage_category_wise = round(num(row, c_occupancy, pct=True), 2)

                    continue
                doc.append("occupancy_category_report", {
                    "category": str(category).strip(),
                    "total_rooms": num(row, c_total),
                    "occupied_rooms": num(row, c_occupied),
                    "vacant_rooms": num(row, c_vacant),
                    "occupancy_": num(row, c_occupancy, pct=True)
                })

    except Exception as e:
        frappe.log_error(frappe.get_traceback(),f"Occupancy Report Import Error: {str(e)}")
    # PM-CM-WR
    try:
        pm_df = pd.read_excel(file_path,sheet_name="PM-CM-WR",header=None)
        doc.set("pm", [])
        doc.set("cm", [])
        doc.set("wr", [])
        doc.set("month_summary", [])
        doc.set("overall_summary", [])

        pmi_idx = None
        cm_idx = None
        wr_idx = None
        overall_pm_idx = None
        overall_cm_idx = None
        overall_wr_idx = None
        overall_idx = None
        overall_summary_idx = None
        for idx, row in pm_df.iterrows():

            first_col = (
                str(row[0]).strip()
                if not pd.isna(row[0])
                else ""
            )

            if first_col.upper() == "PMI" and pmi_idx is None:
                pmi_idx = idx

            elif first_col.upper() == "CM" and cm_idx is None:
                cm_idx = idx

            elif first_col.upper() == "WR" and wr_idx is None:
                wr_idx = idx

            elif first_col.upper() == "OVERALL PM":
                overall_pm_idx = idx

            elif first_col.upper() == "OVERALL CM":
                overall_cm_idx = idx

            elif first_col.upper() == "OVERALL WR":
                overall_wr_idx = idx

            elif first_col.upper() == "OVERALL":
                overall_idx = idx

            values = [
                str(v).strip().upper()
                if not pd.isna(v)
                else ""
                for v in row[:4]
            ]

            if values[:4] == ["MONTH", "PM", "CM", "WR"]:
                overall_summary_idx = idx
        if pmi_idx is not None and cm_idx is not None:

            pm_data = pm_df.iloc[pmi_idx + 2:cm_idx]

            for _, row in pm_data.iterrows():

                subject = row[0]

                if pd.isna(subject):
                    continue

                subject = str(subject).strip()

                if (not subject or subject.upper().startswith("TOTAL") or subject.upper() == "SUBJECT" or subject.upper() == "SCHEDULED" or subject.upper() == "ACCOMPLISHED" or subject.upper() == "COMPLETION RATE"):
                    continue

                doc.append("pm", {
                    "subject": subject,
                    "planned": 0 if pd.isna(row[1]) else float(row[1]),
                    "completed": 0 if pd.isna(row[2]) else float(row[2]),
                    "outstanding": 0 if pd.isna(row[3]) else round(float(row[3]) * 100, 2)
                })
        if cm_idx is not None and wr_idx is not None:

            cm_data = pm_df.iloc[cm_idx + 2:wr_idx]

            for _, row in cm_data.iterrows():

                subject = row[0]

                if pd.isna(subject):
                    continue

                subject = str(subject).strip()

                if (not subject or subject.upper().startswith("TOTAL") or subject.upper() == "SUBJECT" or subject.upper() == "SCHEDULED" or subject.upper() == "ACCOMPLISHED" or subject.upper() == "COMPLETION RATE"):
                    continue

                doc.append("cm", {
                    "subject": subject,
                    "scheduled": 0 if pd.isna(row[1]) else float(row[1]),
                    "accomplished": 0 if pd.isna(row[2]) else float(row[2]),
                    "outstanding": 0 if pd.isna(row[3]) else round(float(row[3]) * 100, 2)
                })
        wr_end = overall_pm_idx if overall_pm_idx is not None else overall_idx
        if wr_idx is not None and wr_end is not None:

            wr_data = pm_df.iloc[wr_idx + 2:wr_end]

            for _, row in wr_data.iterrows():

                subject = row[0]

                if pd.isna(subject):
                    continue

                subject = str(subject).strip()

                if (not subject or subject.upper().startswith("TOTAL") or subject.upper() == "SUBJECT" or subject.upper() == "SCHEDULED" or subject.upper() == "ACCOMPLISHED" or subject.upper() == "COMPLETION RATE"):
                    continue

                doc.append("wr", {
                    "subject": subject,
                    "scheduled": 0 if pd.isna(row[1]) else float(row[1]),
                    "accomplished": 0 if pd.isna(row[2]) else float(row[2]),
                    "outstanding": 0 if pd.isna(row[3]) else round(float(row[3]) * 100, 2)
                })
       
        if overall_pm_idx is not None and overall_cm_idx is not None:
            pm_summary = pm_df.iloc[overall_pm_idx + 2:overall_cm_idx]
            for _, row in pm_summary.iterrows():

                if pd.isna(row[0]):
                    continue

                month = str(row[0]).strip()

                if month.upper() in ["MONTH", "TOTAL"]:
                    continue

                doc.append("month_summary", {
                    "month": month,
                    "type": "PM",
                    "scheduled": 0 if pd.isna(row[1]) else float(row[1]),
                    "accomplished": 0 if pd.isna(row[2]) else float(row[2]),
                    "outstanding": 0 if pd.isna(row[3]) else round(float(row[3]) * 100, 2)
                })
        if overall_cm_idx is not None and overall_wr_idx is not None:
            cm_summary = pm_df.iloc[overall_cm_idx + 2:overall_wr_idx]
            for _, row in cm_summary.iterrows():

                if pd.isna(row[0]):
                    continue

                month = str(row[0]).strip()

                if month.upper() in ["MONTH", "TOTAL"]:
                    continue

                doc.append("month_summary", {
                    "month": month,
                    "type": "CM",
                    "scheduled": 0 if pd.isna(row[1]) else float(row[1]),
                    "accomplished": 0 if pd.isna(row[2]) else float(row[2]),
                    "outstanding": 0 if pd.isna(row[3]) else round(float(row[3]) * 100, 2)
                })
        if overall_wr_idx is not None and overall_summary_idx is not None:
            wr_summary = pm_df.iloc[overall_wr_idx + 2:overall_summary_idx]
            for _, row in wr_summary.iterrows():

                if pd.isna(row[0]):
                    continue

                month = str(row[0]).strip()

                if month.upper() in ["MONTH", "TOTAL"]:
                    continue

                doc.append("month_summary", {
                    "month": month,
                    "type": "WR",
                    "scheduled": 0 if pd.isna(row[1]) else float(row[1]),
                    "accomplished": 0 if pd.isna(row[2]) else float(row[2]),
                    "outstanding": 0 if pd.isna(row[3]) else round(float(row[3]) * 100, 2)
                })
        if overall_summary_idx is not None:
            overall_df = pm_df.iloc[overall_summary_idx + 1:]
            for _, row in overall_df.iterrows():
                if pd.isna(row[0]):
                    continue
                month = str(row[0]).strip()
                if not month:
                    continue
                doc.append("overall_summary", {
                    "month": month,
                    "type": "PM",
                    "accomplished": 0 if pd.isna(row[1]) else float(row[1])
                })

                doc.append("overall_summary", {
                    "month": month,
                    "type": "CM",
                    "accomplished": 0 if pd.isna(row[2]) else float(row[2])
                })

                doc.append("overall_summary", {
                    "month": month,
                    "type": "WR",
                    "accomplished": 0 if pd.isna(row[3]) else float(row[3])
                })
    except Exception as e:
        frappe.log_error(frappe.get_traceback(),f"PM-CM-WR Import Error: {str(e)}")
    # Total Activities
    try:
        ta_df = pd.read_excel(file_path,sheet_name="Total Activites",header=None)
        doc.set("total_activities", [])
        for _, row in ta_df.iterrows():
            if pd.isna(row[0]):
                continue
            activity = str(row[0]).strip()
            total_done = 0 if pd.isna(row[1]) else float(row[1])
            if activity.upper() == "TOTAL ACTIVITIES DONE":
                doc.total_activities_done = total_done
            doc.append("total_activities", {
                "activity": activity,
                "total_activities_done": total_done
            })
    except Exception as e:
        frappe.log_error(frappe.get_traceback(),f"Total Activities Import Error: {str(e)}")
    doc.save(ignore_permissions=True)
    frappe.db.commit()

    return {
        "status": "success",
        "monthly_progress_data": doc.name
    }


@frappe.whitelist()
def get_template():
    from io import BytesIO

    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    title_font = Font(bold=True, size=12)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4472C4")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_side = Side(style="thin", color="7F7F7F")
    header_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    label_font = Font(bold=True)
    label_fill = PatternFill("solid", fgColor="D9E1F2")

    wb = Workbook()
    wb.remove(wb.active)

    def make_sheet(name):
        return wb.create_sheet(title=name)

    def autofit(ws, row, headers, start_col=1):
        lines = 1
        for i, h in enumerate(headers):
            if not h:
                continue
            width = min(max(12, len(str(h)) + 4), 40)
            col = get_column_letter(start_col + i)
            ws.column_dimensions[col].width = max(
                ws.column_dimensions[col].width or 0, width
            )
            lines = max(lines, -(-len(str(h)) // max(int(width * 0.9), 1)))
        ws.row_dimensions[row].height = max(20, 16 * lines + 6)

    def put_title(ws, row, text, col=1):
        ws.cell(row=row, column=col, value=text).font = title_font
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = max(
            ws.column_dimensions[letter].width or 0, min(len(str(text)) + 4, 45)
        )

    def put_note(ws, coord, text):
        c = ws[coord]
        c.value = text
        c.font = Font(italic=True, size=10, color="808080")
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[c.row].height = max(
            ws.row_dimensions[c.row].height or 0, 20
        )

    def put_header(ws, row, headers, start_col=1):
        for i, h in enumerate(headers):
            c = ws.cell(row=row, column=start_col + i, value=h)
            c.font = header_font
            c.fill = header_fill
            c.alignment = header_alignment
            c.border = header_border
        autofit(ws, row, headers, start_col)

    def put_label(ws, row, text):
        c = ws.cell(row=row, column=1, value=text)
        c.font = label_font
        c.fill = label_fill
        ws.column_dimensions["A"].width = max(
            ws.column_dimensions["A"].width or 0, len(text) + 4
        )

    # Resident Satisfaction Survey
    ws = make_sheet("Resident Satisfaction Survey")
    put_title(ws, 1, "Resident Satisfaction Survey")
    put_note(ws, "B1", "Chart proposed: Feedback - Weekwise / Average Feedback Achieved")
    ws.merge_cells("B1:H1")
    put_header(ws, 2, ["Attribute", "Week 1", "Week 2", "Week 3", "Week 4"])

    # Customer Satisfaction Survey
    ws = make_sheet("Customer Satisfaction Survey")
    put_title(ws, 1, "Customer Satisfaction Survey")
    put_note(ws, "C1", "chart proposed: Testimonials from random rooms")
    ws.merge_cells("C1:E1")
    put_header(ws, 2, [
        "Resident Name",
        "Room Number",
        "How do you rate our employees responsiveness with you",
        "How do you rate front desk employee",
        "How do you rate room service employees",
        "How do you rate our employees professionalism with you",
        "How do you rate service quality",
        "How do you rate the cleanliness of the room",
        "How do you rate the Technical Service Quality",
        "Taste of the food",
        "Service",
        "Food quality",
        "Restaurant employees service quality",
        "Laundry Service Quality",
    ])

    # Catering and Laundry Services
    ws = make_sheet("Catering and Laundry Services")
    put_title(ws, 1, "Catering and Laundry Services")
    put_note(ws, "B1", "Chart proposed: Feedback - Weekwise / Total Meal served - count summary / Mealwise Breakup / Laundry count")
    ws.merge_cells("B1:H1")
    put_header(ws, 2, ["Attribute", "Week 1", "Week 2", "Week 3", "Week 4"])
    put_header(ws, 11, ["Meal", "Count"])
    put_label(ws, 12, "Breakfast")
    put_label(ws, 13, "Lunch")
    put_label(ws, 14, "Dinner")
    put_label(ws, 16, "Laundry")

    # Occupancy Report
    ws = make_sheet("Occupancy Report")
    put_title(ws, 1, "Occupancy Report")
    put_note(ws, "C1", "chart proposed: Total Room vs Occupied / Building wise Occupancy / categorywise Occupancy")
    ws.merge_cells("C1:F1")
    put_header(ws, 2, [
        "BUILDING NO",
        "TOTAL ROOMS",
        "OCCUPIED ROOMS",
        "VACANT ROOMS",
        "OCCUPANCY %",
        "Type of Building",
    ])
    put_label(ws, 8, "TOTAL")
    put_header(ws, 10, [
        "Category",
        "Total Rooms",
        "Occupied Rooms",
        "Vacant Rooms",
        "Occupancy %",
    ])
    put_label(ws, 16, "TOTAL")

    # PM-CM-WR
    ws = make_sheet("PM-CM-WR")
    put_label(ws, 1, "PMI")
    put_note(ws, "C1", "chart proposed: Overall Scheduled vs Accomplished / Subjectwise comparison")
    ws.merge_cells("C1:E1")
    put_header(ws, 2, ["Subject", "Planned", "Completed", "Outstanding"])
    put_label(ws, 11, "Total PMI")
    put_label(ws, 12, "Scheduled")
    put_label(ws, 13, "Accomplished")
    put_label(ws, 14, "Completion Rate")
    put_label(ws, 16, "CM")
    put_note(ws, "C16", "chart proposed: Overall Scheduled vs Accomplished / Subjectwise comparison")
    ws.merge_cells("C16:E16")
    put_header(ws, 17, ["Subject", "Scheduled", "Accomplished", "%"])
    put_label(ws, 24, "Total")
    put_label(ws, 25, "Scheduled")
    put_label(ws, 26, "Accomplished")
    put_label(ws, 27, "Completion Rate")
    put_label(ws, 29, "WR")
    put_note(ws, "C29", "chart proposed: Overall Scheduled vs Accomplished / Subjectwise comparison")
    ws.merge_cells("C29:E29")
    put_header(ws, 30, ["Subject", "Scheduled", "Accomplished", "[%]"])
    put_label(ws, 37, "Total")
    put_label(ws, 38, "Scheduled")
    put_label(ws, 39, "Accomplished")
    put_label(ws, 40, "Completion Rate")
    put_label(ws, 42, "Overall PM")
    put_note(ws, "C42", "chart proposed:  Monthwise for PMI / CM / WR and summary of all the 3")
    ws.merge_cells("C42:E42")
    put_header(ws, 43, ["Month", "Scheduled", "Accomplished", "Outstanding"])
    put_label(ws, 50, "Overall CM")
    put_note(ws, "B50", "chart proposed:  Monthwise for PMI / CM / WR and summary of all the 3")
    ws.merge_cells("B50:E50")
    put_header(ws, 51, ["Month", "Scheduled", "Accomplished", "Outstanding"])
    put_label(ws, 58, "Overall WR")
    put_note(ws, "B58", "chart proposed:  Monthwise for PMI / CM / WR and summary of all the 3")
    ws.merge_cells("B58:E58")
    put_header(ws, 59, ["Month", "Scheduled", "Accomplished", "Outstanding"])
    put_header(ws, 66, ["Month", "PM", "CM", "WR"])

    # Total Activites (sheet name kept as-is: importer looks up this exact name)
    ws = make_sheet("Total Activites")
    put_label(ws, 1, "Total Activities Done")

    # Manpower
    ws = make_sheet("Manpower")
    put_label(ws, 1, "Manpower Summary")
    put_note(ws, "C1", "chart proposed: Functionwise Summary and Positionwise")
    ws.merge_cells("C1:E1")
    put_header(ws, 2, ["FUNCTION", "Indicative", "Actual"])
    put_label(ws, 13, "Manpower Detailed")
    put_header(ws, 14, ["FUNCTION", "Indicative", "Actual"])

    # SAUDISATION
    ws = make_sheet("SAUDISATION")
    put_title(ws, 1, "SAUDISATION")
    put_header(ws, 2, ["Functions", "Total", "Saudi Nationals"])

    # Safety
    ws = make_sheet("Safety")
    put_title(ws, 1, "SAFETY")
    put_note(ws, "B1", "chart proposed: Summary  / Casewise")
    ws.merge_cells("B1:C1")
    put_header(ws, 2, ["Injuiry Recorded", "Count"])

    # Energy Management
    ws = make_sheet("Energy Management")
    put_title(ws, 1, "Energy Management")
    put_note(ws, "B1", "chart proposed: Daywise / Overall Consumption")
    put_title(ws, 2, "Daywise Electricity Consumption")
    ws.merge_cells("A2:B2")
    put_header(ws, 3, ["DATE", "Unit Consumed"])

    # Water Consumption (single table: importer reads the first table only)
    ws = make_sheet("Water Consumption")
    put_title(ws, 1, "Water Consumption")
    ws.merge_cells("A1:B1")
    put_header(ws, 2, ["DATE", "RAW  Water", "Sweet Water"])

    # Sewage water & Garbage  (sheet name kept as-is: trailing space required by importer)
    ws = make_sheet("Sewage water & Garbage ")
    put_header(ws, 1, ["DATE", "Sewage Water", "Garbage Collection"])
    put_header(ws, 2, ["", "Consumption(Litres)", "Capacity"])
    ws.merge_cells("A1:A2")

    # Gasoline Consumption
    ws = make_sheet("Gasoline Consumption")
    put_title(ws, 1, "Gasoline Consumption")
    ws.merge_cells("A1:C1")
    put_header(ws, 2, ["DATE", "GASOLINE 91", "Diesel"])

    # Safety & Environment Checklist
    ws = make_sheet("Safety & Environment Checklist")
    put_title(ws, 1, "Safety & Environment Checklist")
    for col, text in [
        (1, "SUMMARY OF SCORES AND PERCENTAGE"),
        (2, "WEIGHTED SCORE"),
        (4, "SATISFACTORY RATING"),
    ]:
        c = ws.cell(row=2, column=col, value=text)
        c.font = header_font
        c.fill = header_fill
        c.alignment = header_alignment
        c.border = header_border
    autofit(ws, 2, ["SUMMARY OF SCORES AND PERCENTAGE", "WEIGHTED SCORE", "", "SATISFACTORY RATING"])
    for col, text in [(2, "APPLICABLE"), (3, "ACTUAL")]:
        c = ws.cell(row=3, column=col, value=text)
        c.font = header_font
        c.fill = header_fill
        c.alignment = header_alignment
        c.border = header_border
    ws.row_dimensions[3].height = 22
    # same merged header layout as the original workbook
    ws.merge_cells("A2:A3")
    ws.merge_cells("B2:C2")
    ws.merge_cells("D2:D3")

    # Training On
    ws = make_sheet("Training On")
    put_title(ws, 1, "Training On")
    ws.merge_cells("A1:D1")
    put_header(ws, 2, ["Training On", "Planned", "Conducted", "Achieved %"])

    # KPI Title
    ws = make_sheet("KPI Title")
    put_header(ws, 1, ["KPI Title", "Performance Measure", "Achieved"])

    for ws in wb.worksheets:
        span = {}
        for mr in ws.merged_cells.ranges:
            span[(mr.min_row, mr.min_col)] = (mr.min_col, mr.max_col)
        for row in ws.iter_rows():
            for cell in row:
                if not (cell.value and cell.font and cell.font.italic):
                    continue
                c0, c1 = span.get((cell.row, cell.column), (cell.column, cell.column))
                width = sum(
                    ws.column_dimensions[get_column_letter(c)].width or 10
                    for c in range(c0, c1 + 1)
                )
                chars_per_line = max(int(width) - 4, 15)
                lines = -(-len(str(cell.value)) // chars_per_line)
                ws.row_dimensions[cell.row].height = max(
                    ws.row_dimensions[cell.row].height or 0, lines * 14 + 6, 20
                )

    out = BytesIO()
    wb.save(out)

    frappe.response["filename"] = "Monthly_Progress_Upload_Template.xlsx"
    frappe.response["filecontent"] = out.getvalue()
    frappe.response["type"] = "binary"