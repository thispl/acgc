// Copyright (c) 2026, divya.p@groupteampro.com and contributors
// For license information, please see license.txt

function show_site_id_error(frm, message, fieldname) {
    frappe.msgprint({
        title: __("Site ID Validation Failed"),
        message: __(message || "Site ID validation failed."),
        indicator: "red",
        clear: true
    });
    if (fieldname) {
        frm.set_value(fieldname, "");
    }
}

function toggle_template_button(frm) {
    var field = frm.fields_dict.site_id;
    if (!field) {
        return;
    }
    var $btn = field.$wrapper.find(".get-template-btn");
    $btn.toggle(!!frm.doc.select_site);
}

function place_template_button(frm) {
    var field = frm.fields_dict.site_id;
    if (!field) {
        return;
    }
    var $wrap = field.$wrapper;
    var $btn = $wrap.find(".get-template-btn");
    var $input = $wrap.find(".control-input-wrapper");
    if (!$btn.length || !$input.length) {
        return;
    }
    $input.css({ display: "flex", "align-items": "center", "flex-wrap": "wrap" });
    $input.find(".control-value").css({ flex: "1 1 auto", "min-width": "0" });
    $btn.css({ display: "", "flex-shrink": "0", "margin-left": "8px" }).appendTo($input);
}

function validate_attached_file(frm, fieldname) {
    fieldname = fieldname || "upload";
    const file_url = frm.doc[fieldname];
    if (!file_url) {
        return;
    }
    const check_key = (frm.doc.select_site || "") + "|" + file_url;
    if (!frm._site_id_check_keys) {
        frm._site_id_check_keys = {};
    }
    if (frm._site_id_check_keys[fieldname] === check_key) {
        return;
    }
    frm._site_id_check_keys[fieldname] = check_key;
    frappe.call({
        method: fieldname === "organization_chart_upload"
            ? "acgc.acgc.doctype.monthly_progress_upload.monthly_progress_upload.validate_org_chart"
            : "acgc.acgc.doctype.monthly_progress_upload.monthly_progress_upload.validate_site_id",
        args: {
            file_url: file_url,
            site: frm.doc.select_site
        },
        freeze: true,
        freeze_message: __("Validating Site ID..."),
        callback(r) {
            if (r.message && r.message.status === "success") {
                frappe.show_alert({
                    message: __(r.message.message),
                    indicator: "green"
                });
            } else {
                frm._site_id_check_keys[fieldname] = null;
                show_site_id_error(frm, r.message && r.message.message, fieldname);
            }
        },
        error() {
            frm._site_id_check_keys[fieldname] = null;
            show_site_id_error(frm, null, fieldname);
        }
    });
}

frappe.ui.form.on("Monthly Progress Upload", {
    refresh(frm) {
        var attach_control = frm.fields_dict.upload;
        if (attach_control) {
            attach_control.save_form = function () {};
        }
        var org_attach_control = frm.fields_dict.organization_chart_upload;
        if (org_attach_control) {
            org_attach_control.save_form = function () {};
        }
        attach_control.$wrapper.find(".get-template-btn").remove();
        var $w = frm.fields_dict.site_id ? frm.fields_dict.site_id.$wrapper : attach_control.$wrapper;
        $w.find(".get-template-btn").remove();
        var $target = $w.find(".control-input-wrapper");
        if (!$target.length) $target = $w.find(".control-input");
        if (!$target.length) $target = $w;
        $('<button class="btn btn-xs btn-default get-template-btn" style="margin-left:8px;">Get Template</button>')
            .appendTo($target)
            .on("click", function (e) {
                e.preventDefault();
                if (!frm.doc.select_site) {
                    frappe.show_alert(
                        {
                            message: __("Please select a Site before downloading the template."),
                            indicator: "red"
                        },
                        10
                    );
                    return;
                }
                if (!frm.doc.posting_date) {
                    frappe.show_alert(
                        {
                            message: __("Please select a Posting Date before downloading the template."),
                            indicator: "red"
                        },
                        10
                    );
                    return;
                }
                fetch(
                    "/api/method/acgc.acgc.doctype.monthly_progress_upload.monthly_progress_upload.get_template?site=" +
                    encodeURIComponent(frm.doc.select_site) +
                    "&posting_date=" +
                    encodeURIComponent(frm.doc.posting_date)
                )
                    .then(function (response) {
                        if (!response.ok) {
                            return response.json().then(function (data) {
                                var message = __("Unable to download the template.");
                                try {
                                    var server_messages = JSON.parse(data._server_messages || "[]");
                                    if (server_messages.length) {
                                        var parsed = JSON.parse(server_messages[0]);
                                        if (parsed.message) {
                                            message = parsed.message;
                                        }
                                    }
                                } catch (err) {}
                                frappe.show_alert({ message: __(message), indicator: "red" }, 10);
                            });
                        }
                        var filename = "MPR_Template.xlsx";
                        var disposition = response.headers.get("Content-Disposition") || "";
                        var match = disposition.match(/filename="?([^";]+)"?/i);
                        if (match && match[1]) {
                            filename = match[1];
                        }
                        return response.blob().then(function (blob) {
                            var url = URL.createObjectURL(blob);
                            var link = document.createElement("a");
                            link.href = url;
                            link.download = filename;
                            document.body.appendChild(link);
                            link.click();
                            link.remove();
                            URL.revokeObjectURL(url);
                        });
                    })
                    .catch(function () {
                        frappe.show_alert(
                            { message: __("Unable to download the template."), indicator: "red" },
                            10
                        );
                    });
            });

        place_template_button(frm);
        toggle_template_button(frm);

        frm.add_custom_button("Upload Data", function () {
            const run_import = () => {
                frappe.call({
                    method: "acgc.acgc.doctype.monthly_progress_upload.monthly_progress_upload.import_monthly_progress",
                    args: {
                        upload_doc: frm.doc.name
                    },
                    freeze: true,
                    freeze_message: "Uploading....",
                    callback(r) {
                        if (r.message && r.message.status === "success") {
                            const month_name = r.message.period_month || "";
                            const yr = r.message.period_year || "";
                            frappe.show_alert({
                                message: month_name ? __("Monthly Progress Data Created/Updated for {0} {1}", [month_name, yr]) : __("Monthly Progress Data Updated"),
                                indicator: "green"
                            });
                            if (r.message.monthly_progress_data) {
                                frappe.set_route("Form", "Monthly Progress Data", r.message.monthly_progress_data);
                            }
                        }
                    }
                });
            };

            const do_upload = () => {
                frappe.call({
                    method: "acgc.acgc.doctype.monthly_progress_upload.monthly_progress_upload.validate_site_id",
                    args: {
                        upload_doc: frm.doc.name
                    },
                    freeze: true,
                    freeze_message: "Validating Site ID...",
                    callback(r) {
                        if (r.message && r.message.status === "success") {
                            frappe.show_alert({
                                message: __(r.message.message),
                                indicator: "green"
                            });
                            run_import();
                        } else {
                            var failed_url = r.message && r.message.file_url;
                            var failed_field = null;
                            if (failed_url === frm.doc.upload) {
                                failed_field = "upload";
                            } else if (failed_url === frm.doc.organization_chart_upload) {
                                failed_field = "organization_chart_upload";
                            }
                            show_site_id_error(frm, r.message && r.message.message, failed_field);
                        }
                    },
                    error() {
                        show_site_id_error(frm);
                    }
                });
            };

            if (frm.is_dirty() || frm.is_new()) {
                frm.save().then(do_upload);
            } else {
                do_upload();
            }
        });
    },

    upload(frm) {
        validate_attached_file(frm, "upload");
    },

    organization_chart_upload(frm) {
        validate_attached_file(frm, "organization_chart_upload");
    },

    select_site(frm) {
        toggle_template_button(frm);
        validate_attached_file(frm, "upload");
        validate_attached_file(frm, "organization_chart_upload");
    }
});
