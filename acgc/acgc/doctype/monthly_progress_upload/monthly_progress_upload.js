// Copyright (c) 2026, divya.p@groupteampro.com and contributors
// For license information, please see license.txt

frappe.ui.form.on("Monthly Progress Upload", {
	refresh(frm) {
        frm.add_custom_button("Upload Data",function(){

            frappe.call({
            method: "acgc.acgc.doctype.monthly_progress_upload.monthly_progress_upload.import_monthly_progress",
            args: {
                upload_doc: frm.doc.name
            },
            freeze:true,
            freeze_message:"Uploading....",
            callback(r) {
                if (r.message) {
                    frappe.msgprint(
                        __("Monthly Progress Data Updated")
                    );
                }
            }
        });
    
        })
    }   
});
