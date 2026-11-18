# Copyright (c) 2026, Divya and contributors
# For license information, please see license.txt

import calendar

import frappe
from frappe.model.document import Document
from frappe.utils import getdate


class MonthlyProgressData(Document):
    # pass
    def validate(self):
        if self.safety_and_environment_checklist:

            total_applicable = 0
            total_actual = 0

            for i in self.safety_and_environment_checklist:
                total_applicable += float(i.applicable or 0)
                total_actual += float(i.actual or 0)

            self.overall_applicable = total_applicable
            self.overall_actual = total_actual

            if total_applicable:
                self.overall_rating = round(
                    (total_actual / total_applicable) * 100,
                    2
                )
            else:
                self.overall_rating = 0

        if self.gasoline_consumption:
            total_gasoline=0
            total_diesel=0
            for i in self.gasoline_consumption:
                total_gasoline+=float(i.gasoline_91)
                total_diesel+=float(i.diesel)
            self.total_gasoline=total_gasoline
            self.total_diesel=total_diesel
        if self.sewage_water_and_garbage:
            total_sewage_water=0
            total_garbage_collection=0
            for i in self.sewage_water_and_garbage:
                total_sewage_water+=float(i.sewage_water)
                total_garbage_collection+=float(i.garbage_collection)
            self.total_sewage_water=total_sewage_water
            self.total_garbage_collection=total_garbage_collection
        if self.water_consumption:
            total_raw_water=0
            total_sweet_water=0
            for i in self.water_consumption:
                total_raw_water+=float(i.raw_water)
                total_sweet_water+=float(i.sweet_water)
            self.total_water_consumption=total_raw_water+total_sweet_water
        if self.energy_management:
            total_energy=0
            for i in self.energy_management:
                total_energy+=float(i.unit_consumed)
            self.total=round(total_energy,0)
        if self.saudisation:
            total_manpower = 0
            total_saudi_nationals = 0
            for i in self.saudisation:
                total_manpower += float(i.total or 0)
                total_saudi_nationals += float(i.saudi_nationals or 0)

            self.total_manpower = total_manpower
            self.total_saudi_nationals = total_saudi_nationals

        if self.organization_chart and (str(self.organization_chart).lower().endswith(".xlsx") or str(self.organization_chart).lower().endswith(".xls")):
            try:
                from acgc.utils.org_chart_generator import generate_org_chart_svg, extract_reporting_period
                file_doc = frappe.get_doc("File", {"file_url": self.organization_chart})
                full_path = file_doc.get_full_path()
                # The dashboard filters by month/year of posting_date, so file this
                # record under the workbook's own reporting period (a month header
                # inside the file, else its file name) rather than the save date.
                period = extract_reporting_period(full_path)
                if period:
                    year, month = period
                    d = getdate(self.posting_date)
                    self.posting_date = d.replace(
                        year=year, month=month,
                        day=min(d.day, calendar.monthrange(year, month)[1]))
                svg_url = generate_org_chart_svg(site=self.site, excel_file_path=full_path, posting_date=self.posting_date)
                self.organization_chart = svg_url
            except Exception as e:
                frappe.log_error(frappe.get_traceback(), f"Org Chart SVG Conversion Error: {str(e)}")