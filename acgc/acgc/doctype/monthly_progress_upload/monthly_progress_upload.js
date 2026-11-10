// Copyright (c) 2026, divya.p@groupteampro.com and contributors
// For license information, please see license.txt

frappe.ui.form.on("Monthly Progress Upload", {
	refresh(frm) {
        var $w = frm.fields_dict.upload.$wrapper;
        $w.find(".get-template-btn").remove();
        var $target = $w.find(".control-input-container .control-input");
        if (!$target.length) $target = $w.find(".control-input");
        if (!$target.length) $target = $w;
        $('<button class="btn btn-xs btn-default get-template-btn" style="margin-left:8px;">Get Template</button>')
            .appendTo($target)
            .on("click", function(e){
                e.preventDefault();
                window.location.href = "/api/method/acgc.acgc.doctype.monthly_progress_upload.monthly_progress_upload.get_template";
            });

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
