/*
 * Moodle JS Content Script
 * Scrape Moodle web pages.
 */


// import "browser_polyfill_mv3.js"
// import {DeepPartial, sleep, throwf, Errorlike} from "shared.js"


namespace MJS {



    export type Page_Data_Base = { moodle_page: Moodle_Page_Data; page: string; dom_submit?: boolean|string; };

    export type Moodle_Page_Data = {
        wwwroot:    string;
        sesskey:    string;
        body_class: string;
        editing:    boolean;
    };


    export type Page_Data =
          page_admin_report_customsql_index_data
        | page_admin_report_customsql_view_data
        | page_backup_backup_data
        | page_backup_backupfilesedit_data
        | page_backup_restore_data
        | page_backup_restorefile_data
        | page_course_editsection_data
        | page_course_index_data
        | page_course_management_data
        | page_course_view_data
        | page_grade_report_grader_index_data
        | page_local_otago_login_data
        | page_login_index_data
        | page_mod_feedback_edit_data
        | page_mod_feedback_use_templ_data
        | page_module_edit_data
        | page_my_index_data
        | Page_Data_Base & { page: ".*"; dom_submit?: boolean|string; };  // some other page



    function moodle_page(): Moodle_Page_Data {
        const logout_dom = document.querySelector<HTMLAnchorElement>("a[href*='logout.php'][role='menuitem']");
        let wwwroot = window.location.origin;
        if ((window.location.pathname.split("/")?.[1] || "").match(/^moodle(?:_\d+)?$/)) {
            wwwroot = wwwroot + '/' + window.location.pathname.split("/")[1];
        }
        return {
            wwwroot:    wwwroot,
            body_class: window.document.body.className!,
            sesskey:    logout_dom ? logout_dom.search.match(/^\?sesskey=(\w+)$/)![1] : "",
            editing:    window.document.body.classList.contains("editing")
        };
    }



    export type block_data = {
        id:         number;
        blockname:  string;
        configdata_unserialised: {
            title?: string;
            text?:  string;
        }
    };


    async function block(): Promise<block_data[]> {
        const blocks_dom = document.querySelectorAll<HTMLElement>("aside.block-region > .block");
        const blocks: block_data[] = [];
        for (const block_dom of Object.values(blocks_dom)) {
            const block_type    = block_dom.getAttribute("data-block")!;
            if (block_type == "adminblock" && (!block_dom.getAttribute("id") || !block_dom.getAttribute("id")!.match(/^inst([0-9]+)$/))) { continue; }
            const block_id      = parseInt(block_dom.getAttribute("id")!.match(/^inst([0-9]+)$/)![1]);
            const block_title_dom = block_dom.querySelector<HTMLHeadingElement>(":scope > div > .card-title, :scope > div.header > div.title > h2");
            const block_text_dom = block_dom.querySelector<HTMLDivElement>(":scope > div > div.card-text.content > div.no-overflow, :scope > div.content > div.no-overflow");
            blocks.push({
                id:         block_id,
                blockname:  block_type,
                configdata_unserialised: {
                    title:  block_title_dom ? block_title_dom.textContent! : undefined,
                    text:   block_text_dom ? block_text_dom.innerHTML : undefined
                }
            });
        }
        return blocks;
    }


    export type page_admin_report_customsql_index_data = Page_Data_Base & {
        page:       "admin-report-customsql-index",
        location?:  { pathname: "/report/customsql/index.php" }
        query_cats: page_admin_report_customsql_category[];
    };

    export type page_admin_report_customsql_category = {
        readonly id:    number;
        name:           string;
        mdl_report_customsql_queries: page_admin_report_customsql_query[];
    };

    export type page_admin_report_customsql_query = {
        readonly id:    number;
        displayname:    string;
    };


    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    async function page_admin_report_customsql_index(_message: DeepPartial<page_admin_report_customsql_index_data>): Promise<page_admin_report_customsql_index_data> {
        const cats_dom = document.querySelectorAll<HTMLDivElement>("#region-main div.csql_category");
        const cats: page_admin_report_customsql_category[] = [];
        for (const cat_dom of Object.values(cats_dom)) {
            const cat_header_dom = cat_dom.querySelector<HTMLAnchorElement>(":scope > h2 > a.categoryname")!;
            const cat_id = parseInt(cat_header_dom.search.match(/\?(?:hide|show)cat=([0-9]+)/)![1]);
            const cat_name = cat_header_dom.textContent!;
            const queries_dom = cat_dom.querySelectorAll<HTMLParagraphElement>(":scope > div.csql_category_reports > p");
            const queries: page_admin_report_customsql_query[] = [];
            for (const query_dom of Object.values(queries_dom)) {
                const query_link_dom = query_dom.querySelector<HTMLAnchorElement>(":scope > a:first-child")!;
                const query_id = parseInt(query_link_dom.search.match(/^\?id=([0-9]+)$/)![1]);
                const query_name = query_link_dom.textContent!;
                queries.push({id: query_id, displayname: query_name});
            }
            cats.push({id: cat_id, name: cat_name, mdl_report_customsql_queries: queries});
        }
        return {moodle_page: moodle_page(), page: "admin-report-customsql-index", query_cats: cats};
    }


    export type page_admin_report_customsql_view_data = Page_Data_Base & {
        page:       "admin-report-customsql-view",
        location?:  { pathname: "/report/customsql/view.php", search: { id: number } }
        query_params?:  {name: string, value: string}[]
        query_results?: page_admin_report_customsql_results;
    };

    export type page_admin_report_customsql_results = {headers: string[], data: string[][]};

    async function page_admin_report_customsql_view(message: DeepPartial<page_admin_report_customsql_view_data>): Promise<page_admin_report_customsql_view_data> {
        let params_out: {name: string, value: string}[]|undefined;
        const params_dom = document.querySelectorAll<HTMLDivElement>("#region-main div[id^='fitem_id_queryparam']");
        if (params_dom.length > 0) {
            params_out = [];
            for (const param_dom of Object.values(params_dom)) {
                const param_out = {name: param_dom.querySelector<HTMLLabelElement>(":scope label")!.textContent!,
                                    value: param_dom.querySelector<HTMLInputElement>(":scope input")!.value};
                if (message.query_params) {
                    const param_in = message.query_params.find((element) => element?.name == param_out.name);
                    if (param_in && param_in.value !== undefined) {
                        param_dom.querySelector<HTMLInputElement>(":scope input")!.value = param_in.value;
                    }
                }
                params_out.push(param_out);
            }
        }

        let results: page_admin_report_customsql_results|undefined;
        const results_dom = document.querySelector<HTMLTableElement>("#region-main table")!;
        if (results_dom) {
            const headers_dom = results_dom.querySelectorAll<HTMLTableHeaderCellElement>(":scope > thead > tr > th");
            const headers: string[] = [];
            for (const header_dom of Object.values(headers_dom)) {
                headers.push(header_dom.textContent!);
            }

            const data_dom = results_dom.querySelectorAll<HTMLTableRowElement>(":scope > tbody > tr");
            const rows: string[][] = [];
            for (const row_dom of Object.values(data_dom)) {
                const row: string[] = [];
                const cells_dom = row_dom.querySelectorAll<HTMLTableDataCellElement>(":scope > td");
                for (const cell_dom of Object.values(cells_dom)) {
                    row.push(cell_dom.textContent!);
                }
                rows.push(row);
            }
            results = {headers: headers, data: rows};
        }

        return {moodle_page: moodle_page(), page: "admin-report-customsql-view", query_params: params_out, query_results: results};
    }


    export type page_backup_backup_base_data = Page_Data_Base & {
        page:       "backup-backup",
        location?:  { pathname: "/backup/backup.php", search: { id: number } },
        stage:      1|2|4|null,
        stage_user: 1|2|3|4|5,
        dom_submit?: "final step"|"next"|"perform backup"|"continue"
    };

    export type page_backup_backup_1_data = page_backup_backup_base_data & {
        stage:      1,
        stage_user: 1,
        include_users: boolean
    };

    export type page_backup_backup_2_data = page_backup_backup_base_data & {
        stage:      2,
        stage_user: 2
    };

    export type page_backup_backup_4_data = page_backup_backup_base_data & {
        stage:      4,
        stage_user: 3,
        backup:    { filename: string },
    };

    export type page_backup_backup_last_data = page_backup_backup_base_data & {
        stage:      null,
        stage_user: 4|5,
        cont_button: boolean
    };

    export type page_backup_backup_data = page_backup_backup_1_data | page_backup_backup_2_data | page_backup_backup_4_data | page_backup_backup_last_data;

    async function page_backup_backup(message: DeepPartial<page_backup_backup_data>): Promise<page_backup_backup_data> {
        const stage_dom = document.querySelector<HTMLInputElement>("#region-main form input[name='stage']");
        const stage: number|null = stage_dom ? parseInt(stage_dom.value) : null;
        const stage_user_dom = document.querySelector<HTMLElement>(
            "#region-main div[role='main'] div#executionprogress:not([style*='display: none;']) div.backup_progress .backup_stage_current,"
            + " #region-main div[role='main'] > div.backup_progress .backup_stage_current")!;
        const stage_user = parseInt(stage_user_dom.textContent!);

        if (message.stage && message.stage !== stage) {
            throw new Error("Page backup backup: Stage mismatch");
        }

        switch (stage) {

            case 1: // Stage 1 for user.
                if (stage_user != 1) { throw new Error("Page backup backup: Stage vs stage user mismatch"); }
                const step1_include_users_dom   = document.querySelector<HTMLInputElement>("#region-main form input#id_setting_root_users[type='checkbox']")!;  // TODO: Not present without appropriate permissions.
                const step1_to_final_step_dom   = document.querySelector<HTMLInputElement>("#region-main form input#id_oneclickbackup[type='submit']")!;
                const step1_next_dom            = document.querySelector<HTMLInputElement>("#region-main form input#id_submitbutton[type='submit']")!;
                if ("include_users" in message && step1_include_users_dom.checked != message.include_users) {
                    step1_include_users_dom.click();
                }
                const include_users = step1_include_users_dom.checked;
                if (message.dom_submit && message.dom_submit == "final step") {
                    step1_to_final_step_dom.click();
                } else if (message.dom_submit && message.dom_submit == "next") {
                    step1_next_dom.click();
                }
                return {moodle_page: moodle_page(), page: "backup-backup", stage: stage, stage_user: stage_user, include_users: include_users};
                break;

            case 2: // Stage 2 for user.
                if (stage_user != 2) { throw new Error("Page backup backup: Stage vs stage user mismatch"); }
                const step2_next_dom = document.querySelector<HTMLInputElement>("#region-main form input#id_submitbutton[type='submit']")!;
                if (message.dom_submit && message.dom_submit == "next") {
                    step2_next_dom.click();
                }
                return {moodle_page: moodle_page(), page: "backup-backup", stage: stage, stage_user: stage_user};
                break;

            case 4: // Stage 3 for user.
                if (stage_user != 3) { throw new Error("Page backup backup: Stage vs stage user mismatch"); }
                const step4_filename = document.querySelector<HTMLInputElement>("#region-main form input#id_setting_root_filename[type='text']")!.value;
                const step4_next_dom = document.querySelector<HTMLInputElement>("#region-main form input#id_submitbutton[type='submit']")!;
                if (message.dom_submit && message.dom_submit == "perform backup") {
                    step4_next_dom.click();
                }
                return {moodle_page: moodle_page(), page: "backup-backup", stage: stage, stage_user: stage_user, backup: {filename: step4_filename}};
                break;

            case null:  // Stages 4 & 5 for user.
                if (stage_user != 4 && stage_user != 5) { throw new Error("Page backup backup: Stage vs stage user mismatch"); }
                const final_step_cont_dom = document.querySelector<HTMLElement>(
                    "#region-main div.continuebutton form [type='submit'], #region-main div.progressbar_container a.btn-primary[href*='restorefile.php']")!;
                if (message.dom_submit && message.dom_submit == "continue") {
                    final_step_cont_dom.click();
                }
                return {moodle_page: moodle_page(), page: "backup-backup", stage: null, stage_user: stage_user, cont_button: final_step_cont_dom ? true : false};
                break;

            default:
                throw new Error("Page backup backup: Stage unrecognised");
        }
    }




    export type page_backup_backupfilesedit_data = Page_Data_Base & {
        page:       "backup-backupfilesedit",
        location?:  { pathname: "/backup/backupfilesedit.php" },
        backups: { filename: string, click?: boolean }[],
        backup?:    { click?: "delete"|"delete_ok" },
        dom_submit?: "save"|"cancel"
    };


    async function page_backup_backupfilesedit(message_in: DeepPartial<page_backup_backupfilesedit_data>): Promise<page_backup_backupfilesedit_data> {

        const backup_filemanager_dom = document.querySelector("#region-main form.mform div.filemanager")!;
        // alert (backup_filemanager_dom.outerHTML);  // TODO: Remove.
        do {
            await sleep(100);
        } while (!backup_filemanager_dom.classList.contains("fm-loaded") || backup_filemanager_dom.querySelector(":scope div.fp-content")!.children.length <= 0);
        // await sleep(200);   // TODO: Check actually loaded?
        const backup_list_dom   = backup_filemanager_dom.querySelector(":scope div.filemanager-container div.fm-content-wrapper div.fp-content")!;
        const backups_dom       = backup_list_dom.querySelectorAll(":scope div.fp-file.fp-hascontextmenu, :scope span.fp-filename-icon.fp-hascontextmenu");
        const save_button_dom   = document.querySelector<HTMLInputElement>("#region-main form.mform input#id_submitbutton[type='submit']")!;
        const delete_button_dom = document.querySelector<HTMLButtonElement>("div.moodle-dialogue div.filemanager form button.fp-file-delete")!;

        const message_out: page_backup_backupfilesedit_data = {moodle_page: moodle_page(), page: "backup-backupfilesedit", backups: []};

        for (const backup_dom of Object.values(backups_dom)) {
            const backup_file_link = backup_dom.querySelector<HTMLAnchorElement>(":scope > a:first-child")!;
            const backup_filename = backup_file_link.querySelector(":scope .fp-filename")!.textContent!;
            if (message_in && message_in.backups) {
                const backup_file_in_index = message_in.backups.findIndex(function(value) { return value.filename == backup_filename; });
                if (backup_file_in_index > -1) {
                    if (message_in.backups[backup_file_in_index].click) {
                        backup_file_link.click();
                        await sleep(100);
                    }
                    message_in.backups.splice(backup_file_in_index, 1);
                }
            }
            message_out.backups.push({filename: backup_filename});

        }
        if (message_in.backups && message_in.backups.length > 0) {
            throw new Error("Backup file not found: " + message_in.backups[0].filename);
        }


        if (message_in.backup && message_in.backup.click == "delete") {
            delete_button_dom.click();
            await sleep(100);
        } else if (message_in.backup && message_in.backup.click == "delete_ok") {
            const delete_ok_button_dom = document.querySelector<HTMLButtonElement>("div.moodle-dialogue div.filemanager button.fp-dlg-butconfirm, div.modal-dialog button.btn-primary[data-action='save']")!;
            delete_ok_button_dom.click();
            await sleep(100);
        }

        if (message_in.dom_submit == "save") {
            save_button_dom.click();
            // await sleep(100);
        }

        return message_out;
    }


    export type page_backup_restore_data =
          page_backup_restore_data_2
        | page_backup_restore_data_4d
        | page_backup_restore_data_4s
        | page_backup_restore_data_8
        | page_backup_restore_data_16
        | page_backup_restore_data_final;

    export type page_backup_restore_data_base = Page_Data_Base & {
        page:       "backup-restore",
        location?:  { pathname: "/backup/restore.php" }
        stage:      number|null,
        stage_user: number;
        // mdl_course?: { template_id?: number }
    };

    export type page_backup_restore_data_2  = page_backup_restore_data_base & { stage: 2, stage_user: 1; dom_submit?: "stage 2 submit" };
    export type page_backup_restore_data_4d = page_backup_restore_data_base & { stage: 4, stage_user: 2; displayed_stage: "Destination", mdl_course_category?: { course_category_id: number, name: string }, dom_submit?: "stage 4 new cat search"|"stage 4 new continue" };
    export type page_backup_restore_data_4s = page_backup_restore_data_base & { stage: 4, stage_user: 3; displayed_stage: "Settings", restore_settings: { users?: boolean }, dom_submit?: "stage 4 settings submit" };
    export type page_backup_restore_data_8  = page_backup_restore_data_base & { stage: 8, stage_user: 4; mdl_course: { fullname: string, shortname: string, startdate: number }, dom_submit?: "stage 8 submit" };
    export type page_backup_restore_data_16 = page_backup_restore_data_base & { stage: 16, stage_user: 5; dom_submit?: "stage 16 submit" };
    export type page_backup_restore_data_final = page_backup_restore_data_base & { stage: null, stage_user: 6|7; mdl_course: { course_id?: number } };

    async function page_backup_restore(message: DeepPartial<page_backup_restore_data>): Promise<page_backup_restore_data> {
        const stage_dom = document.querySelector<HTMLInputElement>("#region-main div form input[name='stage']");
        const stage: number|null = stage_dom ? parseInt(stage_dom.value) : null;
        const stage_user_dom = document.querySelector<HTMLElement>("#region-main div[role='main'] > div.backup_progress .backup_stage_current")!;
        const stage_user = parseInt(stage_user_dom.textContent!);

        if (message.stage && message.stage !== stage) {
            throw new Error("Page backup restore: Stage mismatch");
        }

        switch (stage) {

            case 2: // Stage 1 for user.
                if (stage_user != 1) { throw new Error("Page backup restore: Stage vs stage user mismatch"); }
                const stage_2_submit_dom = document.querySelector<HTMLButtonElement>("#region-main div.backup-restore form [type='submit']")!;
                if (message.dom_submit && message.dom_submit == "stage 2 submit") {
                    stage_2_submit_dom.click();
                }
                return {moodle_page: moodle_page(), page: "backup-restore", stage: stage, stage_user: stage_user};
                break;

            case 4: // Stages 2 & 3 for user.
                if (stage_user != 2 && stage_user != 3) { throw new Error("Page backup restore: Stage vs stage user mismatch"); }

                const displayed_stage: "Destination"|"Settings" = document.querySelector<HTMLSpanElement>("#region-main div.backup_progress span.backup_stage_current")!.textContent!.match(/Destination|Settings/)![0] as "Destination"|"Settings";

                if (displayed_stage == "Destination") {
                    if (stage_user != 2) { throw new Error("Page backup restore: Stage vs stage user mismatch"); }

                    // Destination

                    message = message as DeepPartial<page_backup_restore_data_4d>;

                    const stage_4_new_cat_name_dom = document.querySelector<HTMLInputElement>("#region-main div.backup-course-selector.backup-restore form.mform div.bcs-new-course.backup-section input[name='catsearch'][type='text']")!;
                    const stage_4_new_cat_search_dom = document.querySelector<HTMLInputElement>("#region-main div.backup-course-selector.backup-restore form.mform div.bcs-new-course.backup-section .restore-course-search [type='submit']")!;
                    const stage_4_new_continue_dom = document.querySelector<HTMLInputElement>("#region-main div.backup-course-selector.backup-restore form.mform div.bcs-new-course.backup-section input[value='Continue']")!;
                    if (message.stage == 4 && message.displayed_stage == "Destination" && message.mdl_course_category && message.mdl_course_category.name) {
                        stage_4_new_cat_name_dom.value = message.mdl_course_category.name;
                    }

                    if (message.stage == 4 && message.displayed_stage == "Destination" && message.mdl_course_category && message.mdl_course_category.course_category_id) {
                        const stage_4_new_cat_id_dom = document.querySelector<HTMLInputElement>("#region-main div.backup-course-selector.backup-restore form.mform div.bcs-new-course.backup-section input[name='targetid'][type='radio'][value='" + message.mdl_course_category.course_category_id + "']")!;
                        stage_4_new_cat_id_dom.click();
                    }

                    if (message.dom_submit && message.dom_submit == "stage 4 new cat search") {
                        stage_4_new_cat_search_dom.click();
                    }
                    if (message.dom_submit && message.dom_submit == "stage 4 new continue") {
                        stage_4_new_continue_dom.click();
                    }

                    return { moodle_page: moodle_page(), page: "backup-restore", stage: stage, stage_user: stage_user, displayed_stage: displayed_stage };

                } else {
                    if (stage_user != 3) { throw new Error("Page backup restore: Stage vs stage user mismatch"); }

                    // Settings

                    message = message as DeepPartial<page_backup_restore_data_4s>;

                    const stage_4_settings_users_dom = document.querySelector<HTMLInputElement>("#region-main form.mform fieldset#id_rootsettings input[name='setting_root_users'][type='checkbox']");
                    const stage_4_settings_submit_dom = document.querySelector<HTMLInputElement>("#region-main form.mform input[name='submitbutton'][type='submit']")!;

                    if (message.stage == 4 && message.displayed_stage == "Settings" && message.restore_settings) {
                        if (/*message.restore_settings.hasOwnProperty("users") &&*/ message.restore_settings.users != undefined) {
                            stage_4_settings_users_dom!.checked = message.restore_settings.users;   // TODO: Check
                            stage_4_settings_users_dom!.dispatchEvent(new Event("change"));
                            await sleep(100);
                        }
                    }
                    const message_out_restore_settings = { users: stage_4_settings_users_dom?.checked };

                    if (message.dom_submit && message.dom_submit == "stage 4 settings submit") {
                        stage_4_settings_submit_dom.click();
                    }

                    return { moodle_page: moodle_page(), page: "backup-restore", stage: stage, stage_user: stage_user, displayed_stage: displayed_stage, restore_settings: message_out_restore_settings };

                }


                break;

            case 8: // Stage 4 for user.
                if (stage_user != 4) { throw new Error("Page backup restore: Stage vs stage user mismatch"); }
                message = message as DeepPartial<page_backup_restore_data_8>;
                const course_name_dom           = document.querySelector<HTMLInputElement>("#region-main form.mform fieldset#id_coursesettings input[name^='setting_course_course_fullname'][type='text']")!;
                const course_shortname_dom      = document.querySelector<HTMLInputElement>("#region-main form.mform fieldset#id_coursesettings input[name^='setting_course_course_shortname'][type='text']")!;
                const course_startdate_day_dom  = document.querySelector<HTMLSelectElement>("#region-main form.mform fieldset#id_coursesettings select[name^='setting_course_course_startdate'][name$='[day]']")!;
                const course_startdate_month_dom = document.querySelector<HTMLSelectElement>("#region-main form.mform fieldset#id_coursesettings select[name^='setting_course_course_startdate'][name$='[month]']")!;
                const course_startdate_year_dom = document.querySelector<HTMLSelectElement>("#region-main form.mform fieldset#id_coursesettings select[name^='setting_course_course_startdate'][name$='[year]']")!;
                const submit_dom                = document.querySelector<HTMLInputElement>("#region-main form.mform input[name='submitbutton'][type='submit']")!;

                if (message.stage == 8 && message.mdl_course && message.mdl_course.fullname) {
                    course_name_dom.value = message.mdl_course.fullname;
                }

                if (message.stage == 8 && message.mdl_course && message.mdl_course.shortname) {
                    course_shortname_dom.value = message.mdl_course.shortname;
                }

                if (message.stage == 8 && message.mdl_course && message.mdl_course.startdate) {
                    const startdate = new Date(message.mdl_course.startdate * 1000);
                    course_startdate_day_dom.value  = "1";  // Set the day low initially, to avoid overflow when changing the year or month.
                    course_startdate_year_dom.value =  "" + startdate.getUTCFullYear();
                    course_startdate_month_dom.value =  "" + (startdate.getUTCMonth() + 1); // TODO: Check
                    course_startdate_day_dom.value  =  "" + startdate.getUTCDate();
                }

                const message_out_mdl_course = {
                    fullname: course_name_dom.value,
                    shortname: course_shortname_dom.value,
                    startdate: Date.UTC(
                                    parseInt(course_startdate_year_dom.value),
                                    parseInt(course_startdate_month_dom.value) - 1,
                                    parseInt(course_startdate_day_dom.value)
                                ) / 1000
                };

                if (message.dom_submit && message.dom_submit == "stage 8 submit") {
                    submit_dom.click();
                }

                return {moodle_page: moodle_page(), page: "backup-restore", stage: stage, stage_user: stage_user, mdl_course: message_out_mdl_course};
                break;

            case 16: // Stage 5 for user.
                if (stage_user != 5) { throw new Error("Page backup restore: Stage vs stage user mismatch"); }
                const submit16_dom = document.querySelector<HTMLInputElement>("#region-main form.mform input[name='submitbutton'][type='submit']")!;

                if (message.dom_submit && message.dom_submit == "stage 16 submit") {
                    submit16_dom.click();
                }

                return {moodle_page: moodle_page(), page: "backup-restore", stage: stage, stage_user: stage_user};
                break;

            case null:
                if (stage_user != 6 && stage_user != 7) { throw new Error("Page backup restore: Stage vs stage user mismatch"); }
                const course_id_1_dom = document.querySelector<HTMLInputElement>("#region-main div.continuebutton form input[name='id'][type='hidden']");
                const course_id_2_dom = document.querySelector<HTMLAnchorElement>("#region-main div.progressbar_container p a[href*='view.php']");
                const course_id = course_id_1_dom ? parseInt(course_id_1_dom.value)
                                 : (course_id_2_dom ? parseInt(course_id_2_dom.href.match(/view.php\?id=([0-9]+)/)![1])
                                    : undefined);
                const submitcomplete_dom = document.querySelector<HTMLElement>("#region-main div.continuebutton form [type='submit'], "
                    + "#region-main div.progressbar_container p a[href*='view.php']");
                if (message.dom_submit && message.dom_submit == "stage complete submit") {
                    submitcomplete_dom!.click();
                }
                return {moodle_page: moodle_page(), page: "backup-restore", stage: null, stage_user: stage_user, mdl_course: {course_id: course_id}};
                break;
            default:
                throw new Error("Page backup restore: stage not recognised.");
        }

    }



    export type page_backup_restorefile_data = Page_Data_Base & {
        page: "backup-restorefile",
        location?: {pathname: "/backup/restorefile.php", search: {contextid: number}},
        mdl_course: {course_id?: number, backups: {filename: string, download_url?: string}[]},
        mdl_user: {backups: {filename: string, download_url?: string}[]}
    };

    async function page_backup_restorefile(message: DeepPartial<page_backup_restorefile_data>): Promise<page_backup_restorefile_data> {
        const backup_heading_doms = document.querySelectorAll("#region-main h2, #region-main h3");
        if (backup_heading_doms.length < 3 || backup_heading_doms.length > 4 || backup_heading_doms[1].parentNode != backup_heading_doms[2].parentNode) {
            throw new Error("Backup regions not identified");
        }
        let course_backups_dom = null;
        let user_backups_dom = null;
        let course_manage_button_dom = null;
        let user_manage_button_dom = null;
        for (let backup_region = 1; backup_region < 3; backup_region++) {
            for (let backup_dom = backup_heading_doms[backup_region];
                    backup_dom && (backup_region + 1 >= backup_heading_doms.length || backup_dom != backup_heading_doms[backup_region + 1]);
                    backup_dom = backup_dom.nextElementSibling) {
                if (backup_dom.classList.contains("backup-files-table") || backup_dom.querySelector(":scope .backup-files-table")) {
                    if (backup_region == 1) {
                        course_backups_dom = backup_dom.querySelector(":scope tbody");
                    } else {
                        user_backups_dom = backup_dom.querySelector(":scope tbody");
                    }
                }
                if (backup_dom.classList.contains("singlebutton") || backup_dom.querySelector(":scope .singlebutton")) {
                    if (backup_region == 1) {
                        course_manage_button_dom = backup_dom.querySelector<HTMLButtonElement>(":scope form [type='submit']");
                    } else {
                        user_manage_button_dom = backup_dom.querySelector<HTMLButtonElement>(":scope form [type='submit']");
                    }
                }
            }
        }

        const course_restore_link = course_backups_dom?.querySelector<HTMLAnchorElement>(":scope tr  td.cell.c4 a[href*='&component=backup&filearea=course&']");

        const course_backups: {filename: string, download_url?: string}[] = [];
        if (course_backups_dom && !course_backups_dom.classList.contains("empty")) {
            for (const backup_dom of Object.values(course_backups_dom.querySelectorAll<HTMLTableRowElement>(":scope tr"))) {
                course_backups.push({filename: backup_dom.querySelector(":scope td.cell.c0")!.textContent!, download_url: (backup_dom.querySelector<HTMLAnchorElement>(":scope td.cell.c3 a"))?.href});
            }
        }

        const user_backups: {filename: string, download_url?: string}[] = [];
        if (user_backups_dom && !user_backups_dom.classList.contains("empty")) {
            for (const backup_dom of Object.values(user_backups_dom.querySelectorAll<HTMLTableRowElement>(":scope tr"))) {
                user_backups.push({filename: backup_dom.querySelector(":scope td.cell.c0")!.textContent!, download_url: (backup_dom.querySelector<HTMLAnchorElement>(":scope td.cell.c3 a"))?.href});
            }
        }

        if (message.dom_submit && message.dom_submit == "course_restore") {
            course_restore_link!.click();
        } else if (message.dom_submit && message.dom_submit == "course_manage") {
            course_manage_button_dom!.click();
        } else if (message.dom_submit && message.dom_submit == "user_manage") {
            user_manage_button_dom!.click();
        }
        return {moodle_page: moodle_page(), page: "backup-restorefile", mdl_course: {backups: course_backups}, mdl_user: {backups: user_backups}};
    }





    export type page_course_editsection_data = Page_Data_Base & {
        page:       "course-editsection";
        location?:  { pathname: "/course/editsection.php", search: { id: number } },
        mdl_course?: { course_id: number },
        mdl_course_section: page_course_editsection_section;
    };
    type page_course_editsection_section = {
        course_section_id: number;
        name:       string;
        summary:    string;
        options:    { level?: number; }
    };


    async function page_course_editsection(message: DeepPartial<page_course_editsection_data>): Promise<page_course_editsection_data> {
        // const section_id    = message.sectionid;
        // Start
        const section_in = message.mdl_course_section;

        const section_dom: HTMLFormElement      = document.querySelector<HTMLFormElement>("#region-main form.mform")!;
        // let section: Partial<MDL_Course_Sections> = (message.mdl_course_sections||{});

        // ID
        const section_id_dom: HTMLInputElement  = section_dom.querySelector<HTMLInputElement>(":scope input[name='id']")!;
        const section_out_id = parseInt(section_id_dom.value);

        // Name
        const section_name_dom: HTMLInputElement =    // (section_dom.querySelector<HTMLInputElement>(":scope input[name='name']") ||
                                                        (section_dom.querySelector<HTMLInputElement>(":scope fieldset#id_generalhdr input[name='name[value]']")
                                                        || section_dom.querySelector<HTMLInputElement>(":scope fieldset#id_generalhdr input[name='name']"))!;
        if (section_in && section_in.name != undefined) { // section_in.hasOwnProperty('name')) {
            // const section_name_usedefault_dom: HTMLInputElement|null = section_dom.querySelector(":scope input[name='usedefaultname']");
            const section_name_customise_dom: HTMLInputElement|null  = section_dom.querySelector(":scope fieldset#id_generalhdr input#id_name_customize");
            /*if (section_name_usedefault_dom) {
                section_name_usedefault_dom.checked = false;
                section_name_usedefault_dom.dispatchEvent(new Event("change"));
                await sleep(100);
            }*/
            if (section_name_customise_dom) {
                section_name_customise_dom.checked = true;
                section_name_customise_dom.dispatchEvent(new Event("change"));
                await sleep(100);
            }
            section_name_dom.value = "" + section_in.name;
        }
        const section_out_name = section_name_dom.value;

        // Summary
        const section_summary_dom: HTMLTextAreaElement  = section_dom.querySelector<HTMLTextAreaElement>(":scope fieldset#id_generalhdr textarea[name='summary_editor[text]']")!;
        // const section_summary_item_id: string = ((section_dom.elements.namedItem("summary_editor[itemid]")
        //                                                                            || throwf(new Error("Summary ID not found."))) as HTMLInputElement).value;
        // const user_context = "807782";
        if (section_in && section_in.summary != undefined) {
            section_summary_dom.value = section_in.summary;
        }
        const section_out_summary = section_summary_dom.value;
        // section_out.summary = section_summary_dom.value.replace("https://moodle.op.ac.nz/draftfile.php/"+user_context+"/user/draft/"+section_summary_item_id, "@@PLUGINFILE@@");

        const section_out_x_options: {level?: number} = { };

        // Level
        const section_level_dom = section_dom.querySelector<HTMLSelectElement>(":scope fieldset#id_generalhdr select[name='level']")!;
        if (section_in && section_in.options && section_in.options.level != undefined) {
            section_level_dom.value = "" + section_in.options.level;
        }
        if (section_level_dom) {
            section_out_x_options.level = parseInt(section_level_dom.value);
        }

        // End
        if (section_in && message.dom_submit) { // section_in.x_submit) {
            await sleep(100);
            section_dom.submit();
        }

        const section_out = {
            course_section_id: section_out_id,
            name:       section_out_name,
            summary:    section_out_summary,
            options:  section_out_x_options
        };

        return {moodle_page: moodle_page(), page: "course-editsection", mdl_course_section: section_out};
    }




    export type page_course_index_data = Page_Data_Base & {
        page:       "course-index(-category)?",
        location?:  { pathname: "/course/index.php", search: { categoryid?: number } }
        mdl_course?: { course_id: number },
        mdl_course_category: page_course_index_category;
        dom_expand?: boolean;
    };
    type page_course_index_category = {
        course_category_id: number;
        name:       string;
        description?: string;
        mdl_course_categories: page_course_index_category[];
        mdl_courses: page_course_index_course[];
        more:       boolean;
    };
    type page_course_index_course = {
        course_id:  number;
        fullname:   string;
    };

    async function page_course_index(message: DeepPartial<page_course_index_data>): Promise<page_course_index_data> {

        async function category(category_dom?: HTMLDivElement): Promise<page_course_index_category> {


            async function course(course_dom: HTMLDivElement): Promise<{ course_id: number, fullname: string }> {
                const course_id_out = parseInt(course_dom.dataset.courseid!);
                const course_name_out = course_dom.querySelector<HTMLAnchorElement>(":scope .coursename a")!.text;
                return { course_id: course_id_out, fullname: course_name_out };
            }

            let category_out: page_course_index_category;

            if (!category_dom) {
                // Category ID
                const category_out_match =
                    window.document.body.className!.match(/(?:^|\s)category-(\d+)(?:\s|$)/);

                const category_out_id = category_out_match ? parseInt(category_out_match[1]) : 0;

                if (category_out_id) { // Check properties individually?

                    // Category Name
                    const category_out_name = document.querySelector("div#page-navbar .breadcrumb li:last-of-type")!.textContent!.trim();

                    // Category Description
                    const category_out_description = (document.querySelector("#region-main div.box.generalbox.info .no-overflow") || { innerHTML: "" }).innerHTML;

                    category_out = {
                        course_category_id: category_out_id,
                        name:       category_out_name,
                        description: category_out_description,
                        mdl_course_categories: [],
                        mdl_courses: [],
                        more:       false
                    };

                } else {
                    category_out = { course_category_id: 0, name: "", description: "", mdl_course_categories: [], mdl_courses: [], more: false};
                }
                category_dom = document.querySelector<HTMLDivElement>("#region-main div.course_category_tree")!;
            } else {
                const category_name_dom = category_dom.querySelector<HTMLHeadingElement>(":scope > div.info > .categoryname")!;
                const category_out_id = parseInt(category_dom.dataset.categoryid!);
                const category_out_name = category_name_dom.querySelector<HTMLAnchorElement>(":scope > a")!.text;
                category_out = { course_category_id: category_out_id, name: category_out_name, mdl_course_categories: [], mdl_courses: [], more: false };
                if (message.dom_expand && category_dom.classList.contains("collapsed")) {
                    category_name_dom.click();
                    do {
                        await sleep(200);
                    } while (category_dom.classList.contains("notloaded"));
                }
            }
            const subcategories_out: page_course_index_category[] = [];
            category_out.mdl_course_categories = subcategories_out;
            const courses_out: page_course_index_course[] = [];
            category_out.mdl_courses = courses_out;
            category_out.more = false;
            if (category_dom) {
                for (const subcategory_dom of Object.values(category_dom.querySelectorAll<HTMLDivElement>(":scope > div.content > div.subcategories > div.category"))) {
                    subcategories_out.push(await category(subcategory_dom));
                }

                for (const course_dom of Object.values(category_dom.querySelectorAll<HTMLDivElement>(":scope > div.content > div.courses > div.coursebox"))) {
                    courses_out.push(await course(course_dom));
                }
                // TODO: Check for "View more"? .paging.paging-morelink > a   > 40?
                category_out.more = category_dom.querySelector(":scope > div.content > div.courses > div.paging.paging-morelink") ? true : false;
            }
            return category_out;
        }

        return {
            moodle_page: moodle_page(),
            page: "course-index(-category)?",
            mdl_course_category: await category()
        };

    }



    export type page_course_management_data = Page_Data_Base & {
        page:       "course-management",
        location?:  { pathname: "/course/management.php", search: { categoryid: number, perpage?: number } },
        mdl_course_category: page_course_management_category
        mdl_courses: page_course_management_course[]
    };

    export type page_course_management_category = {
        course_category_id: number;
        name:       string;
        coursecount: number;
        mdl_course_categories: page_course_management_category[];
        checked:    boolean;
        expandable: boolean;
        expanded:   boolean;
    };

    export type page_course_management_course = {
        course_id:  number;
        fullname:   string;
        checked:    boolean;
    };

    async function page_course_management(message_in: DeepPartial<page_course_management_data>): Promise<page_course_management_data> {

        async function category(cat_message_in: DeepPartial<page_course_management_category> | null, dom: HTMLDivElement | HTMLLIElement, top?: boolean): Promise<page_course_management_category> {

            let result: page_course_management_category;
            if (top) {
                result = {
                    course_category_id: 0,
                    name:       "",
                    coursecount: 0,
                    checked:    false,
                    expandable: true,
                    expanded:   true,
                    mdl_course_categories: [],
                };
            } else {
                if (cat_message_in && cat_message_in.hasOwnProperty("expanded") && (cat_message_in.expanded != (dom.getAttribute("aria-expanded") == "true"))) {
                    dom.querySelector<HTMLAnchorElement>(":scope > div > a[data-action='expand']")!.click();
                    do {
                        await sleep(100);
                    } while (!dom.querySelector(":scope > ul"));
                }
                if (cat_message_in && cat_message_in.hasOwnProperty("checked") && (cat_message_in.checked != dom.querySelector<HTMLInputElement>(":scope > div > div input.bulk-action-checkbox")!.checked)) {
                    dom.querySelector<HTMLInputElement>(":scope > div > div input.bulk-action-checkbox")!.click();
                    // TODO: pause?
                }
                result = {
                    course_category_id: parseInt(dom.dataset.id!),
                    name:       dom.querySelector(":scope > div > a.categoryname")!.textContent!,
                    coursecount: parseInt(dom.querySelector(":scope > div > div > span.course-count")!.textContent!),
                    checked:    dom.querySelector<HTMLInputElement>(":scope > div > div input.bulk-action-checkbox")!.checked, // broken?
                    expanded:   dom.getAttribute("aria-expanded") == "true",
                    expandable: (dom.dataset.expandable == "1"),
                    mdl_course_categories: [],
                };
            }
            const subcategories_dom = dom.querySelectorAll<HTMLLIElement>(":scope > ul > li.listitem-category");
            const subcategories_out: page_course_management_category[] = [];
            for (const subcategory_dom of Object.values(subcategories_dom)) {
                const subcategory_id = parseInt(subcategory_dom.dataset.id!);
                const subcategory_in = cat_message_in && cat_message_in.mdl_course_categories && cat_message_in.mdl_course_categories.find(function(value) { return value.course_category_id == subcategory_id; }) || null;
                subcategories_out.push(await category(subcategory_in, subcategory_dom));
            }
            result.mdl_course_categories = subcategories_out;
            return result;
        }

        const course_list_dom = document.querySelectorAll<HTMLLIElement>("#region-main div.course-listing ul li.listitem-course");
        const course_list: page_course_management_course[] = [];
        for (const course_dom of Object.values(course_list_dom)) {
            course_list.push({ course_id: parseInt(course_dom.dataset.id!), fullname: course_dom.querySelector(":scope a.coursename")!.textContent!,
                checked: course_dom.querySelector<HTMLInputElement>(":scope input.bulk-action-checkbox")!.checked });
        }

        return {
            moodle_page: moodle_page(),
            page:       "course-management",
            mdl_course_category: await category(message_in.mdl_course_category || null,
                (document.querySelector<HTMLDivElement>("div.category-listing > div.card-body") || document.querySelector<HTMLDivElement>("div.category-listing"))!, true),
            mdl_courses: course_list
        };
        // Note: div.card-body is present in Moodle >= 3.5.0 Boost theme
    }





    export type page_course_view_data = Page_Data_Base & {
        page: "course-view-[a-z]+",
        location?: { pathname: "/course/view.php", search: { id: number } }
        mdl_course: page_course_view_course;
        mdl_course_section?: page_course_view_course_section;
        mdl_block_instance: block_data[];
    };
    type page_course_view_course = {
        course_id:  number; // Needs editing on.
        // fullname:   string;
        format:     string
        mdl_course_sections: page_course_view_course_section[]
    };
    export type page_course_view_course_section = {
        course_section_id?: number;
        section:    number;
        name:       string;
        visible?:   number;
        summary?:   string;
        mdl_course_modules?: page_course_view_course_module[]
        options?: {level?: number};
    };
    export type page_course_view_course_module = {
        course_module_id: number;
        mdl_module_name: string;
        name:       string;
        intro?:     string;
    };

    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    async function page_course_view(_message: DeepPartial<page_course_view_data>): Promise<page_course_view_data> {

        // Course Start
        const main_dom:         Element = document.querySelector("#region-main")!;
        // const result: Partial<Page_Data> = {};

        const course_out_id =    parseInt((window.document.body.className!.match(/\bcourse-(\d+)\b/)!
                                 )[1]);
        // const course_out_fullname =  document.querySelector<HTMLAnchorElement>(".breadcrumb a[title]")!.title || "";
        const course_out_format =     (window.document.body.className!.match(/\bformat-([a-z]+)\b/)!
                        )[1];

        // Sections
        // const section_container_dom: Element = main_dom.querySelector(":scope .course-content")

        const sections_dom:    NodeListOf<HTMLLIElement> = main_dom.querySelectorAll(":scope ul.topics li.section.main, :scope ul.weeks li.section.main, :scope ul.onetopic li.section.main, :scope ul.sections li.section.main, :scope ul.gtopics li.section.main");
        const single_section_dom = main_dom.querySelector(":scope .single-section li.section.main") ;
        let single_section_out: page_course_view_course_section|undefined;
        let course_out_sections: page_course_view_course_section[] = [];
        for (const section_dom of Object.values(sections_dom)) {



            // Section ID
            let section_out_id: number|undefined;
            // Note: Needs editing on.  Doesn't work for flexsections
            const section_edit_dom = section_dom.querySelector<HTMLAnchorElement>(":scope a.edit.menu-action");
            if (section_edit_dom) {
                const section_id_str    = (section_edit_dom.search.match(/(?:^\?|&)id=(\d+)(?:&|$)/)!
                                        )[1]; // TODO: Use URLSearchParams
                section_out_id = parseInt(section_id_str);
            }
                // Note: Needs editing on.  Doesn't work for onetopic?
                // const section_id_str = (section_dom.querySelector(":scope .content > .sectionname .inplaceeditable")
                //                                                                        ||throwf(new Error("WSC course get content, section name edit not found.")
                //                       ).getAttribute("data-itemid")                    ||throwf(new Error("WSC course get content, section id not found.")
                // section_out.id = parseInt(section_id_str)                                || throwf(new Error("WSC course get content, seciton id 0."));
                // TODO: Try multiple methods?

            // Section Number
            const section_num_str       = (section_dom.id!
                                           .match(/^section-(\d+)$/)!
                                          )[1];
            const section_out_section   = parseInt(section_num_str);  // Note: can be 0

            // Section Name
            const section_out_name_dom  = section_dom.querySelector(":scope .sectionname, :scope .content > .section-title");
            const section_out_name      = section_out_name_dom ? section_out_name_dom.textContent! : "";
                                          // TODO: Remove spurious whitespace.  Note: There may be hidden and visible section names?

            // Section Visible
            const section_out_visible = section_dom.classList.contains("hidden") ? 0 : 1;

            // Section Summary
            const section_summary_container_dom = section_dom.querySelector(":scope > .content > .summary, :scope > .content > .summarytext, :scope > .content > div > .summarytext")!;
            const section_summary_dom  = section_summary_container_dom?.querySelector(":scope .no-overflow");
            const section_out_summary =    section_summary_dom ? section_summary_dom.innerHTML : "";


            // Modules
            let modules_out:      page_course_view_course_module[]|undefined;
            if (section_dom.querySelector(":scope .content > .section") && section_dom.style.display != "none" && section_dom.parentElement?.style.display != "none") {

                const modules_dom: NodeListOf<Element> = (section_dom.querySelector(":scope .content > .section")!  // Note: flexsections can have nested sections.
                                                        ).querySelectorAll(":scope > .activity");
                modules_out = [];

                for (const module_dom of Object.values(modules_dom)) {

                    // Module ID
                    const module_id_str     = (module_dom.id!
                                            .match(/^module-(\d+)$/)!
                                            )[1];
                    const module_out_id = parseInt(module_id_str);

                    // Module Type?
                    const module_modname    = (module_dom.className.match(/(?:^|\s)modtype_([a-z]+)(?:\s|$)/)!
                                            )[1];
                    const module_out_modname =    module_modname;

                    // Module Name
                    const module_out_instance_name =       (module_modname == "label")
                            ? (module_dom.querySelector(":scope .contentwithoutlink") || module_dom.querySelector(":scope .activity-altcontent") || {}
                            ).textContent || ""
                            : (module_dom.querySelector(":scope .instancename") || module_dom.querySelector(":scope .fp-filename")
                            )!.textContent || "";  // TODO: Use innerText to avoid unwanted hidden text with Assignments?
                            // TODO: Check handling of empty strings?
                            // TODO: For folder (to handle inline) if no .instancename, use .fp-filename ???

                    // Module Intro
                    const module_out_instance_intro: string|undefined = (module_modname == "label")  // TODO: Test
                                        ? (module_dom.querySelector(":scope .contentwithoutlink") || module_dom.querySelector(":scope .activity-altcontent") || { innerHTML: ""}
                                        ).innerHTML
                                        : (module_dom.querySelector(":scope .contentafterlink") || module_dom.querySelector(":scope .activity-altcontent") || { innerHTML: undefined }
                                        ).innerHTML;

                    const module_out = {
                        course_module_id:         module_out_id,
                        mdl_module_name:    module_out_modname,
                        name:       module_out_instance_name,
                        intro:      module_out_instance_intro
                    };

                    // Module End
                    modules_out.push(module_out);

                }
            }

            // Section End
            const section_out = {
                course_section_id: section_out_id,
                section:    section_out_section,
                name:       section_out_name,
                summary:    section_out_summary,
                visible:    section_out_visible,
                mdl_course_modules: modules_out
            };


            course_out_sections.push(section_out);
            if (section_dom == single_section_dom) {
                single_section_out = section_out;
            }

        }


        // Get section names from OneTopic tabs, if present (hack)

        if (document.body.classList.contains("format-onetopic") && document.querySelector("#region-main ul.nav.nav-tabs")) {

            // Keep section 0, if shown on all pages.
            if (course_out_sections.length > 0 && course_out_sections[0].section == 0 && course_out_sections[0] != single_section_out) {
                course_out_sections = [ course_out_sections[0] ];
            } else {
                course_out_sections = [];
            }

            // If top-level section, include lower-level section headings?  // TODO: Check
            // TODO: should be if ((include_nested_x || sectionnumber == undefined) ... ?
            let subsections_out: page_course_view_course_section[] = [];
            const tab_bars_dom = document.querySelectorAll<HTMLUListElement>("#region-main ul.nav.nav-tabs");
            // if (document.querySelector("#region-main ul.nav.nav-tabs:nth-child(2) li a.active div.tab_initial")) {
                const subsections_dom = tab_bars_dom[1] ? tab_bars_dom[1].querySelectorAll<HTMLAnchorElement>(":scope li a") : [];
                let is_index = true;
                for (const subsection_dom of Object.values(subsections_dom)) {
                    if (subsection_dom.href && subsection_dom.href.match(/changenumsections.php/)) {
                    } else if (subsection_dom.href && !subsection_dom.classList.contains("active")) {
                        const section_match = subsection_dom.href.match(/\/course\/view.php\?id=(\d+)&section=(\d+)(?:#tabs-tree-start)?$/)
                                                                                        || throwf(new Error("WSC course get content, tab links unrecognised: " + subsection_dom.href));
                        const section_num = parseInt(section_match[2]);
                        subsections_out.push({
                            name:       subsection_dom.title,
                            section:    section_num,
                            options: {level:    is_index ? 0 : 1},
                        });
                    } else {
                        single_section_out!.name = subsection_dom.title;
                        single_section_out!.options = single_section_out!.options || {};
                        single_section_out!.options.level = is_index ? 0 : 1;
                        subsections_out.push(single_section_out!);
                    }
                    is_index = false;
                }
            // }


            // if (sectionnumber == undefined) {
                const other_sections_dom = tab_bars_dom[0].querySelectorAll<HTMLAnchorElement>(":scope li a");
                for (const other_section_dom of Object.values(other_sections_dom)) {

                    if ((other_section_dom).href && other_section_dom.href.match(/changenumsections.php/)) {
                    } else if (other_section_dom.href && !other_section_dom.classList.contains("active")) {
                        const section_match = other_section_dom.href.match(/\/course\/view.php\?id=(\d+)&section=(\d+)(?:#tabs-tree-start)?$/)
                                                                                        || throwf(new Error("WSC course get content, tab links unrecognised: " + other_section_dom.href));
                        const section_num = parseInt(section_match[2]);

                        course_out_sections.push({
                            // course_section_id: 0,
                            name:       other_section_dom.title,
                            // summary:    "",
                            section:    section_num,
                            options: {level:    0},
                            // mdl_course_modules:    [],
                        });
                    } else if (subsections_out.length > 0) {

                        subsections_out[0].name = other_section_dom.title || other_section_dom.text;
                        for (const subsection_out of subsections_out) {
                            course_out_sections.push(subsection_out);
                        }
                        subsections_out = [];

                    } else {
                        single_section_out!.name = other_section_dom.title || other_section_dom.text;
                        single_section_out!.options = single_section_out!.options || {};
                        single_section_out!.options.level = 0;
                        course_out_sections.push(single_section_out!);
                    }
                }
            // }

        } else if (document.body.classList.contains("format-multitopic") && document.body.classList.contains("editing")) {
            const other_sections_dom = document.querySelectorAll<HTMLAnchorElement>("#region-main ul.nav.nav-tabs:first-child li a");
            let section_num = 0;
            for (const other_section_dom of Object.values(other_sections_dom)) {
                if (other_section_dom.href && other_section_dom.href.match(/changenumsections.php/)) { continue; }
                if (!other_section_dom.classList.contains("active")) {
                    const section_match = other_section_dom.href.match(/\/course\/view.php\?id=(\d+)(?:&sectionid=(\d+))?$/)
                                                                                || throwf(new Error("WSC course get content, tab links unrecognised: " + other_section_dom.href));
                    if (section_match[2]) {
                        const section_id = parseInt(section_match[2]);
                        while (course_out_sections[section_num].course_section_id != section_id) { section_num++; }
                    }
                    course_out_sections[section_num].options = course_out_sections[section_num].options || {};
                    course_out_sections[section_num].options!.level = 0;
                } else {
                    let subsections_dom = document.querySelectorAll<HTMLAnchorElement>("#region-main ul.nav.nav-tabs:nth-child(2) li a");
                    if (subsections_dom.length <= 0) {
                        subsections_dom = document.querySelectorAll<HTMLAnchorElement>("#region-main ul.nav.nav-tabs:first-child li a.active");
                    }


                    for (const subsection_dom of Object.values(subsections_dom)) {
                        if (subsection_dom.href && subsection_dom.href.match(/changenumsections.php/)) { continue; }
                        if (!subsection_dom.classList.contains("active")) {
                            const section_match = subsection_dom.href.match(/\/course\/view.php\?id=(\d+)(?:&sectionid=(\d+))?$/)
                                                                                        || throwf(new Error("WSC course get content, tab links unrecognised: " + subsection_dom.href));
                            if (section_match[2]) {
                                const section_id = parseInt(section_match[2]);
                                while (course_out_sections[section_num].course_section_id != section_id) { section_num++; }
                            }
                            course_out_sections[section_num].options = course_out_sections[section_num].options || {};
                            course_out_sections[section_num].options!.level = (subsection_dom == subsections_dom[0]) ? 0 : 1;
                        } else {
                            while (!course_out_sections[section_num].mdl_course_modules) {
                                section_num++;
                            }
                            // course_out_sections[section_num].options = course_out_sections[section_num].options || {};
                            // course_out_sections[section_num].options!.level = (subsection_dom == other_section_dom) ? 0 : 1;
                            // section_num++;
                            let first = true;
                            while (section_num < course_out_sections.length && course_out_sections[section_num].mdl_course_modules) {
                                course_out_sections[section_num].options = course_out_sections[section_num].options || {};
                                course_out_sections[section_num].options!.level = first ? ((subsection_dom == subsections_dom[0]) ? 0 : 1) : 2;
                                if (first) {
                                    single_section_out = course_out_sections[section_num];
                                }
                                section_num++;
                                first = false;
                            }
                        }
                    }
                }
            }
        }

        const course_out: page_course_view_course = {
            course_id:  course_out_id,
            // fullname:   course_out_fullname,
            format:     course_out_format,
            mdl_course_sections: course_out_sections
        };

        return {moodle_page: moodle_page(), page: "course-view-[a-z]+", mdl_course: course_out, mdl_course_section: single_section_out, mdl_block_instance: await block()};
    }



    export type page_grade_report_grader_index_data = Page_Data_Base & {
        page:       "grade-report-grader-index",
        location?:  { pathname: "/grade/report/grader/index.php", search: {id: number} },
        grades_table_as_text: string
    };

    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    async function page_grade_report_grader_index(_message: DeepPartial<page_grade_report_grader_index_data>): Promise<page_grade_report_grader_index_data> {
        // alert("starting page_local_otago_login");
        const grades_table_dom = document.querySelector<HTMLTableElement>("table#user-grades.gradereport-grader-table")!;
        let found_heading: boolean = false;
        let table_as_text: string = "";
        for (const row_dom of Object.values(grades_table_dom.rows)) {
            if (!found_heading) {
                if (!row_dom.classList.contains("heading")) { continue; }
                found_heading = true;
            }
            if (row_dom.classList.contains("lastrow")) { break; }
            // let col_num: number = 0;
            for (const cell_dom of Object.values(row_dom.cells)) {
                table_as_text = table_as_text + cell_dom.textContent;
                for (let col_count = 0; col_count < cell_dom.colSpan; col_count++) {
                    // col_num++;
                    table_as_text = table_as_text + "\t";
                }
            }
            table_as_text = table_as_text + "\n";
        }
        return {moodle_page: moodle_page(), page: "grade-report-grader-index", grades_table_as_text: table_as_text};
    }




    export type page_local_otago_login_data = Page_Data_Base & {
        page:       "local-otago-login",
        location?:  { pathname: "/local/otago/login.php", search: Record<string, string> },
        dom_submit?: "staff_students"|"other_users"
    };

    async function page_local_otago_login(message: DeepPartial<page_local_otago_login_data>): Promise<page_local_otago_login_data> {
        // alert("starting page_local_otago_login");
        const staff_students_dom    = document.querySelector<HTMLAnchorElement>("div.loginactions a[href*='/auth/saml2']")!;
        const other_users_dom       = document.querySelector<HTMLAnchorElement>("div.loginactions a[href*='/login/index']")!;
        if (message.dom_submit == "staff_students") {
            staff_students_dom.click();
        } else if (message.dom_submit == "other_users") {
            other_users_dom.click();
        }
        return {moodle_page: moodle_page(), page: "local-otago-login"};
    }



    export type page_login_index_data = Page_Data_Base & {
        page:       "login-index",
        location?:  { pathname: "/login/index.php", search: Record<string, string> },
        mdl_user?:   page_login_index_user,
        dom_submit?: "log_in"
    };

    type page_login_index_user = {
        username:   string;
        password:   string;
    };

    async function page_login_index(message: DeepPartial<page_login_index_data>): Promise<page_login_index_data> {
        const username_dom  = document.querySelector<HTMLInputElement>("input#username")!;
        const password_dom  = document.querySelector<HTMLInputElement>("input#password")!;
        const log_in_dom    = document.querySelector("#loginbtn") as HTMLAnchorElement;
        if (message.mdl_user) {
            username_dom.value = message.mdl_user.username!;
            await sleep(100);
            password_dom.value = message.mdl_user.password!;
            await sleep(100);
        }
        if (message.dom_submit == "log_in") {
            log_in_dom.click();
        }
        return {moodle_page: moodle_page(), page: "login-index"};
    }



    export type page_mod_feedback_edit_data = Page_Data_Base & {
        page: "mod-feedback-edit",
        location?: { pathname: "/mod/feedback/edit.php", search: {id: number, do_show: "edit"|"templates" } },
        mdl_course_module?: { course_module_id?: number, mdl_feedback_template_id?: number; }
    };

    async function page_mod_feedback_edit(message: DeepPartial<page_mod_feedback_edit_data>): Promise<page_mod_feedback_edit_data> {
       const template_id_dom = document.querySelector<HTMLSelectElement>("#region-main form.mform[action='use_templ.php'] select#id_templateid")!;
       if (message && message.mdl_course_module && message.mdl_course_module
            && message.mdl_course_module.hasOwnProperty("mdl_feedback_template_id")) {
            template_id_dom.value = "" + message.mdl_course_module.mdl_feedback_template_id;
            template_id_dom.dispatchEvent(new Event("change"));

       }
       return { moodle_page: moodle_page(), page: "mod-feedback-edit" };
    }


    export type page_mod_feedback_use_templ_data = Page_Data_Base & {
        page: "mod-feedback-use_templ";
        location?: {pathname: "/mod/feedback/use_templ.php"}
        // mdl_course_modules: {x_submit: boolean;};
        // dom_submit: boolean
    };

    async function page_mod_feedback_use_templ(message: DeepPartial<page_mod_feedback_use_templ_data>): Promise<page_mod_feedback_use_templ_data> {
       const submit_dom = document.querySelector<HTMLInputElement>("#region-main form.mform:not(.feedback_form) input#id_submitbutton")!;
       if (message && message.dom_submit) { // message.mdl_course_modules.x_submit) {
            submit_dom.click();
       }
       return {moodle_page: moodle_page(), page: "mod-feedback-use_templ"};
    }





    export type page_module_edit_data = Page_Data_Base & {
        page:       "mod-[a-z]+-mod",
        location?:  { pathname: "/course/modedit.php", search: { update: number } },
        mdl_course?: { course_id: number }
        mdl_course_module: page_module_edit_module
    };
    type page_module_edit_module = {
        course_module_id: number;
        activity_id: number;
        course:     number;
        section:    number;
        mdl_module_name: string;
        name:   string;
        intro:  string;
    };

    async function page_module_edit(message: DeepPartial<page_module_edit_data>): Promise<page_module_edit_data> {

        // Module Start
        const module_in = message.mdl_course_module;
        // const cmid = message.cmid;
        const module_dom: HTMLFormElement = document.querySelector<HTMLFormElement>("form.mform[action='modedit.php']")!;

        // Module ID
        const module_id_dom         = module_dom.querySelector<HTMLInputElement>("input[name='coursemodule']")!;
        const module_out_id         = parseInt(module_id_dom.value);

        // Module Instance ID
        const module_instance_dom   = module_dom.querySelector<HTMLInputElement>("input[name='instance']")!;
        const module_out_instance   = parseInt(module_instance_dom.value);

        // Module Course
        const module_course_dom     = module_dom.querySelector<HTMLInputElement>("input[name='course']")!;
        const module_out_course     = parseInt(module_course_dom.value);

        // Module Section
        const module_out_section    = parseInt(document.querySelector<HTMLInputElement>("form.mform input[name='section'][type='hidden']")!.value);

        // Module ModName
        const module_modname_dom    = module_dom.querySelector<HTMLInputElement>("input[name='modulename']")!;
        const module_out_modname    = module_modname_dom.value;

        // Module Intro/Description
        const module_description_dom = module_dom.querySelector<HTMLTextAreaElement>("textarea[name='introeditor[text]']")!;
        if (module_in && module_in.intro != undefined) {
            module_description_dom.value = module_in.intro;
        }
        const module_out_instance_intro = module_description_dom.value;

        // Module Name
        const module_name_dom = module_dom.querySelector<HTMLInputElement>("input[name='name']")!;
        // TODO: For label, instead of name field, use introeditor[text] field (without markup)?

        if (module_in && module_in.name != undefined) {
            module_name_dom.value = module_in.name;
        }

        const module_out_instance_name = module_name_dom ? module_name_dom.value : "";


        // Module Completion
        const module_completion_dom = module_dom.querySelector<HTMLSelectElement>("select[name='completion']");
        const module_completion: number = module_completion_dom ? parseInt(module_completion_dom.value) : 0;
        if (module_completion == 0 || module_completion == 1 || module_completion == 2) {  }
        else                                                                        { throw new Error("WSC course get module, completion value unexpected."); }
        // module_out_completion = module_completion;

        // For assignments
        const module_assignsubmission_file_enabled_x_dom = module_dom.querySelector<HTMLInputElement>("input[name='assignsubmission_file_enabled']");
        const module_assignsubmission_onlinetext_enabled_x_dom = module_dom.querySelector<HTMLInputElement>("input[name='assignsubmission_onlinetext_enabled']");

        // TODO: Add these
        // const module_completionview_x_dom:     Element|RadioNodeList|null = module_dom.elements.namedItem("completionview");
        // const module_completionusegrade_x_dom: Element|RadioNodeList|null = module_dom.elements.namedItem("completionusegrade");
        // const module_completionsubmit_x_dom:   Element|RadioNodeList|null = module_dom.elements.namedItem("completionsubmit");

        if (   (    module_id_dom          instanceof HTMLInputElement  && module_id_dom.type         == "hidden"     )
            && (    module_instance_dom    instanceof HTMLInputElement  && module_instance_dom.type   == "hidden"     )
            && (    module_course_dom      instanceof HTMLInputElement  && module_course_dom.type     == "hidden"     )
            && (    module_modname_dom     instanceof HTMLInputElement  && module_modname_dom.type    == "hidden"     )
            && (    module_description_dom instanceof HTMLTextAreaElement                                              ) ) {  }
        else                                                                        { throw new Error("WSC course get module, field type unexpected."); }

        if (   (      !module_name_dom
                || (   module_name_dom                                  instanceof HTMLInputElement
                    && module_name_dom.type                                  == "text"    ))
            &&  (     !module_completion_dom
                || (   module_completion_dom                           instanceof HTMLSelectElement
                    && module_completion_dom.type                            == "select-one")
                || (   module_completion_dom                           instanceof HTMLInputElement
                    && module_completion_dom.type                            == "hidden"  ))
            && (      !module_assignsubmission_file_enabled_x_dom
                || (   module_assignsubmission_file_enabled_x_dom       instanceof HTMLInputElement
                    && module_assignsubmission_file_enabled_x_dom.type       == "checkbox"))
            && (      !module_assignsubmission_onlinetext_enabled_x_dom
                || (   module_assignsubmission_onlinetext_enabled_x_dom instanceof HTMLInputElement
                    && module_assignsubmission_onlinetext_enabled_x_dom.type == "checkbox")))
        {  } else                                                             { throw new Error("WSC course get module, optional field type unexpected."); }

        // if (!module_completionview_x_dom     || module_completionview_x_dom     instanceof HTMLInputElement) {  }
        // else                                                                     { throw new Error("In module, couldn't get completion view."); }
        // if (!module_completionusegrade_x_dom || module_completionusegrade_x_dom instanceof HTMLInputElement) {  }
        // else                                                                     { throw new Error("In module, couldn't get completion grade."); }
        // if (!module_completionsubmit_x_dom   || module_completionsubmit_x_dom   instanceof HTMLInputElement) {  }
        //                                                                          { throw new Error("In module, couldn't get completion submit."); }


        /*
        const module_assignsubmission_file_enabled_x:       0|1|undefined   = module_assignsubmission_file_enabled_x_dom
                                                                              ? (module_assignsubmission_file_enabled_x_dom.checked ? 1 : 0)       : undefined;
        const module_assignsubmission_onlinetext_enabled_x: 0|1|undefined   = module_assignsubmission_onlinetext_enabled_x_dom
                                                                              ? (module_assignsubmission_onlinetext_enabled_x_dom.checked ? 1 : 0) : undefined;
        */
        // TODO: Fix to handle checkboxes appropriately (as above)
        // const module_completionview_x:     number|undefined = module_completionview_x_dom     ? parseInt(module_completionview_x_dom.value)     : undefined;
        // const module_completionusegrade_x: number|undefined = module_completionusegrade_x_dom ? parseInt(module_completionusegrade_x_dom.value) : undefined;
        // const module_completionsubmit_x:   number|undefined = module_completionsubmit_x_dom   ? parseInt(module_completionsubmit_x_dom.value)   : undefined;


        {
            // assignsubmission_file_enabled_x:        module_assignsubmission_file_enabled_x,
            // assignsubmission_onlinetext_enabled_x:  module_assignsubmission_onlinetext_enabled_x,
            // completionview_x:                    module_completionview_x,
            // completionusegrade_x:                module_completionusegrade_x,
            // completionsubmit_x:                  module_completionsubmit_x

        }

        if (module_in && message.dom_submit) { // module_in.x_submit) {
            await sleep(100);
            module_dom.submit();
        }

        const module_out: page_module_edit_module = {
            course_module_id: module_out_id,
            activity_id:   module_out_instance,
            course:     module_out_course,
            section:    module_out_section,
            mdl_module_name: module_out_modname,
            name: module_out_instance_name,
            intro: module_out_instance_intro
        };

        return {moodle_page: moodle_page(), page: "mod-[a-z]+-mod", mdl_course_module: module_out};
    }



    export type page_my_index_data = Page_Data_Base & {
        page:       "my-index",
        location?:  { pathname: "/my/index.php", search: Record<string, string> },
    };

    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    async function page_my_index(_message: DeepPartial<page_my_index_data>): Promise<page_my_index_data> {
        return {moodle_page: moodle_page(), page: "my-index"};
    }




    async function page_onMessage(message: DeepPartial<Page_Data>, sender?: browser.runtime.MessageSender): Promise<Page_Data> {
        if (sender && sender.tab !== undefined) { throw new Error("Unexpected message"); }
        return await page_get_set(message);
    }

    async function page_get_set(message: DeepPartial<Page_Data>): Promise<Page_Data> {

        await sleep(100);

        const error_message_dom = document.querySelector("div.errorbox p.errormessage") ?? document.querySelector("div.alert.alert-danger");
        if (error_message_dom) {
            throw new Error(error_message_dom.textContent!);
        }

        // const message = message_in;

        const body_id = window.document.body.id;
        if (message.page && !body_id.match(RegExp("^page-" + message.page + "$"))) { throw new Error("Unexpected page"); }

        let result: Page_Data;

        switch (body_id) {
            case "page-admin-report-customsql-index":
                result = await page_admin_report_customsql_index(message as DeepPartial<page_admin_report_customsql_index_data>);
                break;
            case "page-admin-report-customsql-view":
                result = await page_admin_report_customsql_view(message as DeepPartial<page_admin_report_customsql_view_data>);
                break;
            case "page-backup-backup":
                result = await page_backup_backup(message as DeepPartial<page_backup_backup_data>);
                break;
            case "page-backup-backupfilesedit":
                result = await page_backup_backupfilesedit(message as DeepPartial<page_backup_backupfilesedit_data>);
                break;
            case "page-backup-restore":
                result = await page_backup_restore(message as DeepPartial<page_backup_restore_data>);
                break;
            case "page-backup-restorefile":
                result = await page_backup_restorefile(message as DeepPartial<page_backup_restorefile_data>);
                break;
            case "page-course-editsection":
                result = await page_course_editsection(message as DeepPartial<page_course_editsection_data>);
                break;
            case "page-course-index":
            case "page-course-index-category":
                result = await page_course_index(message as DeepPartial<page_course_index_data>);
                break;
            case "page-course-management":
                result = await page_course_management(message as DeepPartial<page_course_management_data>);
                break;
            case "page-course-view-onetopic":
            case "page-course-view-multitopic":
            case "page-course-view-topics":
            case "page-course-view-grid":
            case "page-course-view-singleactivity":
            case "page-course-view-social":
            case "page-course-view-weeks":
                result = await page_course_view(message as DeepPartial<page_course_view_data>);
                break;
            case "page-grade-report-grader-index":
                result = await page_grade_report_grader_index(message as DeepPartial<page_grade_report_grader_index_data>);
                break;
            case "page-local-otago-login":
                result = await page_local_otago_login(message as DeepPartial<page_local_otago_login_data>);
                break;
            case "page-login-index":
                result = await page_login_index(message as DeepPartial<page_login_index_data>);
                break;
            case "page-mod-feedback-edit":
                result = await page_mod_feedback_edit(message as DeepPartial<page_mod_feedback_edit_data>);
                break;
            case "page-mod-feedback-use_templ":
                result = await page_mod_feedback_use_templ(message as DeepPartial<page_mod_feedback_use_templ_data>);
                break;
            case "page-mod-assign-mod":
            case "page-mod-assignment-mod":
            case "page-mod-book-mod":
            case "page-mod-chat-mod":
            case "page-mod-choice-mod":
            case "page-mod-data-mod":
            case "page-mod-feedback-mod":
            case "page-mod-folder-mod":
            case "page-mod-forum-mod":
            case "page-mod-glossary-mod":
            // case "page-mod-imscp-mod":
            case "page-mod-journal-mod":
            case "page-mod-label-mod":
            case "page-mod-lesson-mod":
            // case "page-mod-lti-mod":
            case "page-mod-page-mod":
            // case "page-mod-questionnaire-mod":
            case "page-mod-quiz-mod":
            case "page-mod-resource-mod":
            // case "page-mod-scorm-mod":
            // case "page-mod-survey-mod":
            case "page-mod-url-mod":
            case "page-mod-wiki-mod":
            case "page-mod-workshop-mod":
                result = await page_module_edit(message as DeepPartial<page_module_edit_data>);
                break;
            case "page-my-index":
                result = await page_my_index(message as DeepPartial<page_my_index_data>);
                break;
            default:
                result = {moodle_page: moodle_page(), page: ".*"};
                break;
        }


        return result;
    }



    export async function page_init(): Promise<void>/*{status: boolean}*/ {
        browser.runtime.onMessage.addListener(page_onMessage);
        // return {status: true};
        // return c_on_call({});
        let message: Page_Data|Errorlike;
        try {
            message = await page_onMessage({});
        } catch (e: Error) {
            message = {name: "Error", message: e.message, fileName: e.fileName, lineNumber: e.lineNumber};
        }
        void browser.runtime.sendMessage(message);
    }



}


void MJS.page_init();
