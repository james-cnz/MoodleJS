
namespace MJS {



    export type Page_Data_Base = {page: string};

    export type Page_Data =
          page_backup_backup_data
        | page_backup_restore_data
        | page_backup_backupfilesedit_data
        | page_backup_restorefile_data
        | page_course_editsection_data
        | page_course_index_data
        | page_course_view_data
        | page_mod_feedback_edit_data
        | page_mod_feedback_use_templ_data
        | page_module_edit_data
        | {page: "*"};  // some other page


    export type Page_Data_In_Base = {
        dom_submit?:         boolean|string;
    };

    export type Page_Data_In = ({page: string} & DeepPartial<Page_Data> | {}) & Page_Data_In_Base;


    export type Page_Data_Out_Base = {
        moodle_page:        Moodle_Page;
    };

    export type Page_Data_Out = Page_Data & Page_Data_Out_Base;


    export type Moodle_Page = {
        wwwroot:    string;
        sesskey:    string;
        body_id:    string;
        body_class: string;
    };


    export type page_backup_backup_data = Page_Data_Base & {
        page: "backup-backup",
        location?: {pathname: "/backup/backup.php", search: {id: number}}
        stage: 1|2|4|null
        backup?: {filename: string}
        dom_submit?: "final step"|"next"|"perform backup"|"continue"
    };


    async function page_backup_backup(message: Page_Data_In_Base & DeepPartial<page_backup_backup_data>): Promise<page_backup_backup_data> {
        const stage_dom = document.querySelector("#region-main div form input[name='stage']") as HTMLInputElement|null;
        const stage: number|null = stage_dom ? parseInt(stage_dom.value) : null;

        if (message.stage && message.stage !== stage) {
            throw new Error("Page backup backup: Stage mismatch");
        }

        switch (stage) {

            case 1:
                const step1_to_final_step_dom   = document.querySelector("#region-main form input#id_oneclickbackup[type='submit']")    as HTMLButtonElement;
                const step1_next_dom            = document.querySelector("#region-main form input#id_submitbutton[type='submit']")      as HTMLButtonElement;
                if (message.dom_submit && message.dom_submit == "final step") {
                    step1_to_final_step_dom.click();
                } else if (message.dom_submit && message.dom_submit == "next") {
                    step1_next_dom.click();
                }
                return {page: "backup-backup", stage: stage};
                break;

            case 2:
                const step2_next_dom = document.querySelector("#region-main form input#id_submitbutton[type='submit']") as HTMLButtonElement;
                if (message.dom_submit && message.dom_submit == "next") {
                    step2_next_dom.click();
                }
                return {page: "backup-backup", stage: stage};
                break;

            case 4:
                const step4_filename = (document.querySelector("#region-main form input#id_setting_root_filename[type='text']") as HTMLInputElement).value;
                const step4_next_dom = document.querySelector("#region-main form input#id_submitbutton[type='submit']") as HTMLButtonElement;
                if (message.dom_submit && message.dom_submit == "perform backup") {
                    step4_next_dom.click();
                }
                return {page: "backup-backup", stage: stage, backup: {filename: step4_filename}};
                break;

            case null:
                const final_step_cont_dom = document.querySelector("#region-main div.continuebutton form button[type='submit']") as HTMLButtonElement;
                if (message.dom_submit && message.dom_submit == "continue") {
                    final_step_cont_dom.click();
                }
                return {page: "backup-backup", stage: null};
                break;

            default:
                throw new Error("Page backup backup: Stage unrecognised");
        }
    }


    export type page_course_index_data = Page_Data_Base & {
        page: "course-index(-category)?",
        location?: {pathname: "/course/index.php", search: {categoryid?: number}}
        mdl_course?: {id: number},
        mdl_course_categories: page_course_index_category;
        dom_expand?: boolean;
    };
    type page_course_index_category = {
        id:         number;
        name:       string;
        description?: string;
        mdl_course_categories: page_course_index_category[];
        mdl_course: page_course_index_course[];
        more:       boolean;
    };
    type page_course_index_course = {
        id:         number;
        fullname:   string;
    };

    async function page_course_index(message: Page_Data_In_Base & DeepPartial<page_course_index_data>): Promise<page_course_index_data> {

        async function category(category_dom?: HTMLDivElement): Promise<page_course_index_category> {


            async function course(course_dom: HTMLDivElement): Promise<{id: number, fullname: string}> {
                const course_id_out = parseInt(course_dom.getAttribute("data-courseid") as string);
                const course_name_out = (course_dom.querySelector(":scope .coursename a") as HTMLAnchorElement).text;
                return {id: course_id_out, fullname: course_name_out};
            }

            let category_out: page_course_index_category;

            if (!category_dom) {
                // Category ID
                const category_out_match =
                    (window.document.body.getAttribute("class")   || throwf(new Error("WSC category get displayed, body class not found."))
                    ).match(/(?:^|\s)category-(\d+)(?:\s|$)/);

                const category_out_id = category_out_match ? parseInt(category_out_match[1]) : 0;

                if (category_out_id) { // Check properties individually?

                    // Category Name
                    const breadcrumbs_dom = window.document.querySelectorAll(":root div#page-navbar .breadcrumb li");  // :last-child or :last-of-type
                    if (breadcrumbs_dom.length > 0) { /*OK*/ } else                            { throw new Error("WSC category get displayed, breadcrumbs not found"); }
                    const breadcrumb_last_dom = breadcrumbs_dom.item(breadcrumbs_dom.length - 1);
                    const category_out_name =  (breadcrumb_last_dom.querySelector(":scope a")     || throwf(new Error("WSC category get displayed, breadcrumb link not found."))
                                            ).textContent                                       || throwf(new Error("WSC category get displayed, name not found."));

                    // Category Description
                    const category_out_description = (window.document.querySelector(":root #region-main div.box.generalbox.info .no-overflow") || { innerHTML: "" }).innerHTML;

                    category_out = {
                        id:         category_out_id,
                        name:       category_out_name,
                        description: category_out_description,
                        mdl_course_categories: [],
                        mdl_course: [],
                        more:       false
                    };

                } else {
                    // category_out_name = "";
                    // category_out_description = "";
                    category_out = { id: 0, name: "", description: "", mdl_course_categories: [], mdl_course: [], more: false};
                }
                category_dom = document.querySelector(".course_category_tree") as HTMLDivElement;
            } else {
                const category_link_dom = category_dom.querySelector(":scope > .info > .categoryname > a") as HTMLAnchorElement;
                const category_out_id = parseInt(category_dom.getAttribute("data-categoryid") as string);
                const category_out_name = category_link_dom.text;
                category_out = {id: category_out_id, name: category_out_name, mdl_course_categories: [], mdl_course: [], more: false};
                if (message.dom_expand && category_dom.classList.contains("collapsed")) {
                    (category_dom.querySelector(":scope > .info > h3.categoryname, :scope > .info > h4.categoryname") as HTMLHeadingElement).click();
                    do {
                        await sleep(200);
                    } while (category_dom.classList.contains("notloaded"));
                }
            }
            const subcategories_out: page_course_index_category[] = [];
            category_out.mdl_course_categories = subcategories_out;
            for (const subcategory_dom of Object.values(category_dom.querySelectorAll(":scope > .content > .subcategories > div.category") as NodeListOf<HTMLDivElement>)) {
                subcategories_out.push(await category(subcategory_dom));
            }
            const courses_out: page_course_index_course[] = [];
            category_out.mdl_course = courses_out;
            for (const course_dom of Object.values(category_dom.querySelectorAll(":scope > .content > .courses > div.coursebox") as NodeListOf<HTMLDivElement>)) {
                courses_out.push(await course(course_dom));
            }
            // TODO: Check for "View more"? .paging.paging-morelink > a   > 40?
            category_out.more = category_dom.querySelector(":scope > .content > .courses > div.paging.paging-morelink") ? true : false;
            return category_out;
        }

        return {
            page: "course-index(-category)?",
            mdl_course_categories: await category()
        };

    }



    export type page_backup_restorefile_data = Page_Data_Base & {
        page: "backup-restorefile",
        location?: {pathname: "/backup/restorefile.php", search: {contextid: number}},
        mdl_course: {id?: number, backups: {filename: string, download_url: string}[]}
    };

    async function page_backup_restorefile(message: Page_Data_In_Base & DeepPartial<page_backup_restorefile_data>): Promise<page_backup_restorefile_data> {
        const course_backups_dom = document.querySelector("table.backup-files-table tbody") as HTMLTableElement;
        // const download_link = document.querySelector(".backup-files-table .c3 a") as HTMLAnchorElement;
        const restore_link = document.querySelector("#region-main table.backup-files-table.generaltable  tbody tr  td.cell.c4.lastcol a[href*='&component=backup&filearea=course&']") as HTMLAnchorElement;
        const manage_button_dom = document.querySelector("section#region-main div.singlebutton form button[type='submit']") as HTMLButtonElement;

        const backups: {filename: string, download_url: string}[] = [];
        if (!course_backups_dom.classList.contains("empty")) {
            for (const backup_dom of Object.values(course_backups_dom.querySelectorAll("tr") as NodeListOf<HTMLTableRowElement>)) {
                backups.push({filename: backup_dom.querySelector("td.cell.c0").textContent, download_url: (backup_dom.querySelector("td.cell.c3 a") as HTMLAnchorElement).href});
            }
        }

        // if (message.dom_submit && message.dom_submit == "download") {
        //    await browser.downloads.download({url: download_link.href, saveAs: false});
        // }
        if (message.dom_submit && message.dom_submit == "restore") {
            restore_link.click();
        } else if (message.dom_submit && message.dom_submit == "manage") {
            manage_button_dom.click();
        }
        return {page: "backup-restorefile", mdl_course: {backups: backups}};
    }


    export type page_backup_backupfilesedit_data = Page_Data_Base & {
        page: "backup-backupfilesedit",
        location?: {pathname: "/backup/backupfilesedit.php"},
        mdl_course: { backups: {filename: string, click?: boolean}[] },
        backup?: { click?: "delete"|"delete_ok"},
        dom_submit?: "save"|"cancel"
    };


    async function page_backup_backupfilesedit(message_in: Page_Data_In_Base & DeepPartial<page_backup_backupfilesedit_data>): Promise<page_backup_backupfilesedit_data> {

        const backup_filemanager_dom = document.querySelector("section#region-main form#mform1 div.filemanager");
        do {
            await sleep(100);
        } while (!backup_filemanager_dom.classList.contains("fm-loaded"));
        await sleep(100);
        const backup_list_dom = document.querySelector("section#region-main form#mform1 div.filemanager div.filemanager-container div.fm-content-wrapper div.fp-content");
        const backups_dom = backup_list_dom.querySelectorAll(".fp-file.fp-hascontextmenu, .fp-filename-icon.fp-hascontextmenu");
        const save_button_dom = document.querySelector("input#id_submitbutton[type='submit']") as HTMLInputElement;
        const delete_button_dom = document.querySelector("button.fp-file-delete") as HTMLButtonElement;

        const message_out: page_backup_backupfilesedit_data = {page: "backup-backupfilesedit", mdl_course: { backups: []}};

        for (const backup_dom of Object.values(backups_dom)) {
            const backup_file_link = backup_dom.querySelector("a");
            const backup_filename = backup_file_link.querySelector(".fp-filename").textContent;
            const backup_file_in_index = (message_in && message_in.mdl_course && message_in.mdl_course.backups) ?
                                            message_in.mdl_course.backups.findIndex(function(value) { return value.filename == backup_filename; })
                                            : -1;
            if (backup_file_in_index > -1) {
                if (message_in.mdl_course.backups[backup_file_in_index].click) {
                    backup_file_link.click();
                    await sleep(100);
                }
                message_in.mdl_course.backups.splice(backup_file_in_index, 1);
            }
            message_out.mdl_course.backups.push({filename: backup_filename});

        }
        if (message_in.mdl_course && message_in.mdl_course.backups && message_in.mdl_course.backups.length > 0) {
            throw new Error("Backup file not found: " + message_in.mdl_course.backups[0].filename);
        }


        if (message_in.backup && message_in.backup.click == "delete") {
            delete_button_dom.click();
            await sleep(100);
        } else if (message_in.backup && message_in.backup.click == "delete_ok") {
            const delete_ok_button_dom = document.querySelector("button.fp-dlg-butconfirm") as HTMLButtonElement;
            delete_ok_button_dom.click();
            await sleep(100);
        }

        if (message_in.dom_submit == "save") {
            save_button_dom.click();
            // await sleep(100);
        }

        return message_out;
    }


    export type page_backup_restore_data = Page_Data_Base & {
        page: "backup-restore",
        location?: {pathname: "/backup/restore.php"}
        stage: 2|4|8|16|null,
        mdl_course?: {template_id?: number}
    } & (
        {stage: 2, dom_submit?: "stage 2 submit"}
        | { stage: 4, mdl_course_categories?: {id: number, name: string}, restore_settings: {users: boolean}, dom_submit?: "stage 4 new cat search"|"stage 4 new continue"|"stage 4 settings submit"}
        | { stage: 8, mdl_course: {fullname: string, shortname: string, startdate: number}, dom_submit?: "stage 8 submit"}
        | { stage: 16, dom_submit?: "stage 16 submit" }
        | { stage: null, mdl_course: {id: number}}
    );

    async function page_backup_restore(message: Page_Data_In_Base & DeepPartial<page_backup_restore_data>): Promise<page_backup_restore_data> {
        const stage_dom = document.querySelector("#region-main div form input[name='stage']") as HTMLInputElement;
        const stage: number|null = stage_dom ? parseInt(stage_dom.value) : null;

        if (message.stage && message.stage !== stage) {
            throw new Error("Page backup restore: Stage mismatch");
        }

        switch (stage) {

            case 2:
                const stage_2_submit_dom = document.querySelector("#region-main div.backup-restore form [type='submit']") as HTMLButtonElement;
                if (message.dom_submit && message.dom_submit == "stage 2 submit") {
                    stage_2_submit_dom.click();
                }
                return {page: "backup-restore", stage: stage};
                break;

            case 4:

                // Destination

                const stage_4_new_cat_name_dom = document.querySelector("#region-main div.backup-course-selector.backup-restore form.mform input[name='catsearch'][type='text']") as HTMLInputElement;
                const stage_4_new_cat_search_dom = document.querySelector("#region-main div.backup-course-selector.backup-restore form.mform input[name='searchcourses'][type='submit']") as HTMLInputElement;
                const stage_4_new_continue_dom = document.querySelector("#region-main div.backup-course-selector.backup-restore form.mform input[value='Continue']") as HTMLInputElement;
                if (message.stage == 4 && message.mdl_course_categories && message.mdl_course_categories.name) {
                    stage_4_new_cat_name_dom.value = message.mdl_course_categories.name;
                }

                if (message.stage == 4 && message.mdl_course_categories && message.mdl_course_categories.id) {
                    const stage_4_new_cat_id_dom = document.querySelector("#region-main div.backup-course-selector.backup-restore form.mform input[name='targetid'][type='radio'][value='" + message.mdl_course_categories.id + "']") as HTMLInputElement;
                    stage_4_new_cat_id_dom.click();
                }

                if (message.dom_submit && message.dom_submit == "stage 4 new cat search") {
                    stage_4_new_cat_search_dom.click();
                }
                if (message.dom_submit && message.dom_submit == "stage 4 new continue") {
                    stage_4_new_continue_dom.click();
                }


                // Settings

                const stage_4_settings_users_dom = document.querySelector("#region-main form#mform1.mform fieldset#id_rootsettings input[name='setting_root_users'][type='checkbox']") as HTMLInputElement;
                const stage_4_settings_submit_dom = document.querySelector("#region-main form#mform1.mform input[name='submitbutton'][type='submit']") as HTMLInputElement;

                if (message.stage == 4 && message.restore_settings) {
                    if (/*message.restore_settings.hasOwnProperty("users") &&*/ message.restore_settings.users != undefined) {
                        stage_4_settings_users_dom.checked = message.restore_settings.users;  // TODO: Check
                        stage_4_settings_users_dom.dispatchEvent(new Event("change"));
                        await sleep(100);
                    }
                }
                const message_out_restore_settings = stage_4_settings_users_dom ? { users: stage_4_settings_users_dom.checked } : null;

                if (message.dom_submit && message.dom_submit == "stage 4 settings submit") {
                    stage_4_settings_submit_dom.click();
                }

                return {page: "backup-restore", stage: stage, restore_settings: message_out_restore_settings};
                break;

            case 8:
                const course_name_dom           = document.querySelector("#region-main form#mform2.mform fieldset#id_coursesettings input[name^='setting_course_course_fullname'][type='text']") as HTMLInputElement;
                const course_shortname_dom      = document.querySelector("#region-main form#mform2.mform fieldset#id_coursesettings input[name^='setting_course_course_shortname'][type='text']") as HTMLInputElement;
                const course_startdate_day_dom  = document.querySelector("#region-main form#mform2.mform fieldset#id_coursesettings select[name^='setting_course_course_startdate'][name$='[day]']") as HTMLSelectElement;
                const course_startdate_month_dom = document.querySelector("#region-main form#mform2.mform fieldset#id_coursesettings select[name^='setting_course_course_startdate'][name$='[month]']") as HTMLSelectElement;
                const course_startdate_year_dom = document.querySelector("#region-main form#mform2.mform fieldset#id_coursesettings select[name^='setting_course_course_startdate'][name$='[year]']") as HTMLSelectElement;
                const submit_dom                = document.querySelector("#region-main form#mform2.mform input[name='submitbutton'][type='submit']") as HTMLInputElement;

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

                return {page: "backup-restore", stage: stage, mdl_course: message_out_mdl_course};
                break;

            case 16:
                const submit16_dom = document.querySelector("#region-main form#mform2.mform input[name='submitbutton'][type='submit']") as HTMLInputElement;

                if (message.dom_submit && message.dom_submit == "stage 16 submit") {
                    submit16_dom.click();
                }

                return {page: "backup-restore", stage: stage};
                break;

            case null:
                const course_id_dom = document.querySelector("#region-main form input[name='id'][type='hidden']") as HTMLInputElement;
                const course_id = parseInt(course_id_dom.value);
                const submitcomplete_dom = document.querySelector("#region-main form [type='submit']") as HTMLButtonElement;
                if (message.dom_submit && message.dom_submit == "stage complete submit") {
                    submitcomplete_dom.click();
                }
                return {page: "backup-restore", stage: null, mdl_course: {id: course_id}};
                break;
            default:
                throw new Error("Page backup restore: stage not recognised.");
        }

    }



    export type page_course_view_data = Page_Data_Base & {
        page: "course-view-[a-z]+",
        location?: {pathname: "/course/view.php", search: {id: number}}
        mdl_course: page_course_view_course;
        mdl_course_sections: page_course_view_course_sections;
    };
    type page_course_view_course = {
        id:         number; // Needs editing on.
        fullname:   string;
        format:     string
        mdl_course_sections: page_course_view_course_sections[]
    };
    type page_course_view_course_sections = {
        id?:        number;
        section:    number;
        name:       string;
        visible?:   number;
        summary?:   string;
        mdl_course_modules?: page_course_view_course_modules[]
        x_options?: {level?: number};
    };
    export type page_course_view_course_modules = {
        id:         number;
        mdl_modules_name: string;
        mdl_course_module_instance: {
            name:       string;
            intro:      string;
        }
    };

    async function page_course_view(_message: Page_Data_In_Base & DeepPartial<page_course_view_data>): Promise<page_course_view_data> {

        // Course Start
        const main_dom:        Element             = window.document.querySelector(":root #region-main")
                                                                                    || throwf(new Error("WSC course get content, main region not found."));
        // const result: Partial<Page_Data> = {};

        const course_out_id =    parseInt((window.document.body.getAttribute("class").match(/\bcourse-(\d+)\b/) || throwf(new Error("WS course get displayed, course id not found."))
                                 )[1]);
        const course_out_fullname =  (window.document.querySelector(":root .breadcrumb a[title]") as HTMLAnchorElement).getAttribute("title") || "";
        const course_out_format =     (window.document.body.getAttribute("class").match(/\bformat-([a-z]+)\b/)    || throwf(new Error("WS course get displayed, course format not found."))
                        )[1];

        // Sections
        // const section_container_dom: Element = main_dom.querySelector(":scope .course-content")

        const sections_dom:    NodeListOf<Element> = main_dom.querySelectorAll(":scope li.main");
        const single_section_dom = main_dom.querySelector(":scope .single-section .main") ;
        let single_section_out: page_course_view_course_sections|undefined;
        let course_out_sections: page_course_view_course_sections[] = [];
        for (const section_dom of Object.values(sections_dom)) {



            // Section ID
            let section_out_id: number|undefined;
            // Note: Needs editing on.  Doesn't work for flexsections
            const section_edit_dom = section_dom.querySelector(":scope a.edit.menu-action") as HTMLAnchorElement|null;
            if (section_edit_dom) {
                const section_id_str    = ((section_edit_dom
                                        ).search.match(/(?:^\?|&)id=(\d+)(?:&|$)/)   || throwf(new Error("WSC course get content, section id not found."))
                                        )[1]; // TODO: Use URLSearchParams
                section_out_id = parseInt(section_id_str);
            }
                // Note: Needs editing on.  Doesn't work for onetopic?
                // const section_id_str = (section_dom.querySelector(":scope > .content > .sectionname .inplaceeditable")
                //                                                                        ||throwf(new Error("WSC course get content, section name edit not found.")
                //                       ).getAttribute("data-itemid")                    ||throwf(new Error("WSC course get content, section id not found.")
                // section_out.id = parseInt(section_id_str)                                || throwf(new Error("WSC course get content, seciton id 0."));
                // TODO: Try multiple methods?

            // Section Number
            const section_num_str       = ((section_dom.getAttribute("id")         || throwf(new Error("WSC course get content, section num not found."))
                                           ).match(/^section-(\d+)$/)               || throwf(new Error("WSC course get content, section num not recognised."))
                                          )[1];
            const section_out_section =    parseInt(section_num_str);  // Note: can be 0

            // Section Name
            const section_out_name =       (section_dom.querySelector(":scope > .content > .sectionname") || throwf(new Error("WSC course get content, section name not found."))
                                          ).textContent                                           || throwf(new Error("WSC course get content, section name text not found."));
                                          // TODO: Remove spurious whitespace.  Note: There may be hidden and visible section names?

            // Section Visible
            const section_out_visible = section_dom.classList.contains("hidden") ? 0 : 1;

            // Section Summary
            const section_summary_container_dom = section_dom.querySelector(":scope > .content > .summary")
                                                                                    || throwf(new Error("WSC course get content, section summary container not found"));
            const section_summary_dom  = section_summary_container_dom.querySelector(":scope .no-overflow");
            const section_out_summary =    section_summary_dom ? section_summary_dom.innerHTML : "";


            // Modules
            let modules_out:      page_course_view_course_modules[]|undefined;
            if (section_dom.querySelector(":scope > .content > .section")) {

                const modules_dom: NodeListOf<Element> = (section_dom.querySelector(":scope > .content > .section")  // Note: flexsections can have nested sections.
                                                                                        || throwf(new Error("WSC course get content, section content not found."))
                                                        ).querySelectorAll(":scope .activity");
                modules_out = [];

                for (const module_dom of Object.values(modules_dom)) {

                    // Module ID
                    const module_id_str     = ((module_dom.getAttribute("id")          || throwf(new Error("WSC course get content, mod ID not found."))
                                            ).match(/^module-(\d+)$/)                || throwf(new Error("WSC course get content, mod ID not recognised."))
                                            )[1];
                    const module_out_id = parseInt(module_id_str) || throwf(new Error("WSC course get content, mod ID 0"));

                    // Module Type?
                    const module_modname    = (module_dom.className.match(/(?:^|\s)modtype_([a-z]+)(?:\s|$)/)
                                                                                        || throwf(new Error("WSC course get content, modname not found."))
                                            )[1];
                    const module_out_modname =    module_modname;

                    // Module Name
                    const module_out_instance_name =       (module_modname == "label")
                            ? (module_dom.querySelector(":scope .contentwithoutlink") || throwf(new Error("WSC course get content, label name not found."))
                            ).textContent || ""
                            : (module_dom.querySelector(":scope .instancename") || throwf(new Error("WSC course get content, name not found."))
                            ).textContent || "";  // TODO: Use innerText to avoid unwanted hidden text with Assignments?
                            // TODO: Check handling of empty strings?
                            // TODO: For folder (to handle inline) if no .instancename, use .fp-filename ???

                    // Module Intro
                    const module_out_instance_intro: string|undefined = (module_modname == "label")  // TODO: Test
                                        ? (module_dom.querySelector(":scope .contentwithoutlink") || throwf(new Error("WSC course get content, label description not found."))
                                        ).innerHTML
                                        : (module_dom.querySelector(":scope .contentafterlink") || { innerHTML: undefined }
                                        ).innerHTML;

                    const module_out = {
                        id:         module_out_id,
                        mdl_modules_name:    module_out_modname,
                        mdl_course_module_instance: {
                            name:       module_out_instance_name,
                            intro:      module_out_instance_intro
                        }
                    };

                    // Module End
                    modules_out.push(module_out);

                }
            }

            // Section End
            const section_out = {
                id:         section_out_id,
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


        // Get section names from OneTopic tabs (hack)

        if (document.body.classList.contains("format-onetopic")) {
            course_out_sections = [];

            // If top-level section, include lower-level section headings?  // TODO: Check
            // TODO: should be if ((include_nested_x || sectionnumber == undefined) ... ?
            let subsections_out: page_course_view_course_sections[] = [];
            // if (document.querySelector("#region-main ul.nav.nav-tabs:nth-child(2) li a.active div.tab_initial")) {
                const subsections_dom = document.querySelectorAll(":root #region-main ul.nav.nav-tabs:nth-child(2) li a") as NodeListOf<HTMLAnchorElement>;
                let is_index = true;
                for (const subsection_dom of Object.values(subsections_dom)) {
                    if (subsection_dom.href) {
                        const section_match = subsection_dom.href.match(/^(https?:\/\/[a-z\-.]+)\/course\/view.php\?id=(\d+)&section=(\d+)$/)
                                                                                        || throwf(new Error("WSC course get content, tab links unrecognised: " + subsection_dom.href));
                        const section_num = parseInt(section_match[3]);
                        subsections_out.push({
                            name:       subsection_dom.title,
                            section:    section_num,
                            x_options: {level:    is_index ? 0 : 1},
                        });
                    } else {
                        single_section_out.x_options = single_section_out.x_options || {};
                        single_section_out.x_options.level = is_index ? 0 : 1;
                        subsections_out.push(single_section_out);
                    }
                    is_index = false;
                }
            // }


            // if (sectionnumber == undefined) {
                const other_sections_dom = document.querySelectorAll(":root #region-main ul.nav.nav-tabs:first-child li a") as NodeListOf<HTMLAnchorElement>;
                for (const other_section_dom of Object.values(other_sections_dom)) {

                    if ((other_section_dom).href && other_section_dom.href.match(/changenumsections.php/)) {
                    } else if (other_section_dom.href) {
                        const section_match = other_section_dom.href.match(/^(https?:\/\/[a-z\-.]+)\/course\/view.php\?id=(\d+)&section=(\d+)$/)
                                                                                        || throwf(new Error("WSC course get content, tab links unrecognised: " + other_section_dom.href));
                        const section_num = parseInt(section_match[3]);

                        course_out_sections.push({
                            id:         0,
                            name:       other_section_dom.title,
                            summary:    "",
                            section:    section_num,
                            x_options: {level:    0},
                            mdl_course_modules:    [],
                        });  // TODO: Any better way to deal with these missing values? (Maybe not.)
                    } else if (subsections_out.length > 0) {

                        for (const subsection_out of subsections_out) {
                            course_out_sections.push(subsection_out);
                        }
                        subsections_out = [];

                    } else {
                        single_section_out.x_options = single_section_out.x_options || {};
                        single_section_out.x_options.level = 0;
                        course_out_sections.push(single_section_out);
                    }
                }
            // }

        }

        const course_out: page_course_view_course = {
            id: course_out_id,
            fullname: course_out_fullname,
            format: course_out_format,
            mdl_course_sections: course_out_sections
        };

        return {page: "course-view-[a-z]+", mdl_course: course_out, mdl_course_sections: single_section_out};
    }


    /*
    export type page_backup_backup_data = Partial<Page_Data> & {
        page: "backup-backup";

    }

    async function page_backup_backup(message: Partial<Page_Data>): Promise<page_backup_backup_data> {

    }
    */


    export type page_course_editsection_data = Page_Data_Base & {
        page: "course-editsection";
        location?: {pathname: "/course/editsection.php", search: {id: number}},
        mdl_course?: {id: number},
        mdl_course_sections: page_course_editsection_section;
    };
    type page_course_editsection_section = {
        id:         number;
        name:       string;
        summary:    string;
        x_options:  { level?: number; }
    };


    async function page_course_editsection(message: Page_Data_In_Base & DeepPartial<page_course_editsection_data>): Promise<page_course_editsection_data> {
        // const section_id    = message.sectionid;
        // Start
        const section_in = message.mdl_course_sections;

        const section_dom:         HTMLFormElement         = window.document.querySelector(":root form#mform1") as HTMLFormElement;
        // let section: Partial<MDL_Course_Sections> = (message.mdl_course_sections||{});

        // ID
        const section_id_dom:      HTMLInputElement  = section_dom.querySelector(":scope input[name='id']") as HTMLInputElement;
        // section.id = parseInt(section_id_dom.value);
        const section_out_id = parseInt(section_id_dom.value);

        // Name
        const section_name_dom:    HTMLInputElement  =    (section_dom.querySelector("input[name='name']")
                                                            || section_dom.querySelector("input[name='name[value]']")) as HTMLInputElement;
        if (section_in && section_in.name != undefined) { // section_in.hasOwnProperty('name')) {
            const section_name_usedefault_dom:  HTMLInputElement|null = section_dom.querySelector("input[name='usedefaultname']");
            const section_name_customise_dom: HTMLInputElement|null = section_dom.querySelector("input#id_name_customize");
            if (section_name_usedefault_dom) {
                section_name_usedefault_dom.checked = false;
                section_name_usedefault_dom.dispatchEvent(new Event("change"));
                await sleep(100);
            }
            if (section_name_customise_dom) {
                section_name_customise_dom.checked = true;
                section_name_customise_dom.dispatchEvent(new Event("change"));
                await sleep(100);
            }
            section_name_dom.value = "" + section_in.name;
        }
        const section_out_name = section_name_dom.value;

        // Summary
        const section_summary_dom: HTMLTextAreaElement  = section_dom.querySelector("textarea[name='summary_editor[text]']")
                                                                                    || throwf(new Error("WSC course get section, summary not found."));
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
        const section_level_dom = section_dom.querySelector("select[name='level']") as HTMLSelectElement;
        if (section_in && section_in.x_options && section_in.x_options.level != undefined) {
            section_level_dom.value = "" + section_in.x_options.level;
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
            id:         section_out_id,
            name:       section_out_name,
            summary:    section_out_summary,
            x_options:  section_out_x_options
        };

        return {page: "course-editsection", mdl_course_sections: section_out};
    }


    export type page_module_edit_data = Page_Data_Base & {
        page: "mod-[a-z]+-mod",
        location?: {pathname: "/course/modedit.php", search: {update: number}},
        mdl_course?: {id: number}
        mdl_course_modules: page_module_edit_module
    };
    type page_module_edit_module = {
        id:         number;
        instance:   number;
        course:     number;
        section:    number;
        mdl_modules_name: string;
        mdl_course_module_instance: Partial<{
            name: string;
            intro: string;
        }>
    };

    async function page_module_edit(message: Page_Data_In_Base & DeepPartial<page_module_edit_data>): Promise<page_module_edit_data> {

        // Module Start
        const module_in = message.mdl_course_modules;
        // const cmid = message.cmid;
        const module_dom:              HTMLFormElement         = window.document.querySelector(":root form#mform1")  as HTMLFormElement;

        // Module ID
        const module_id_dom  = module_dom.elements.namedItem("coursemodule") as HTMLInputElement
                                                                                    || throwf(new Error("WSC course get module, ID not found."));
        const module_out_id =             parseInt(module_id_dom.value);

        // Module Instance ID
        const module_instance_dom  = module_dom.elements.namedItem("instance") as HTMLInputElement
                                                                                    || throwf(new Error("WSC course get module, instance ID not found."));
        const module_out_instance = parseInt(module_instance_dom.value); //                    || throwf(new Error("WSC course get module, instance ID not recognised"));
        // const module_out__instance = {id: module_out.instance};

        // Module Course
        const module_course_dom  = module_dom.elements.namedItem("course") as HTMLInputElement
                                                                                    || throwf(new Error("WSC course get module, course ID not found."));
        const module_out_course = parseInt(module_course_dom.value)                      || throwf(new Error("WSC course get module, course ID not recognised"));

        // Module Section
        const module_out_section = parseInt(((window.document.querySelector(":root form#mform1 input[name='section'][type='hidden']")
        || throwf(new Error("WSC course get module, section num not found.")) ) as HTMLInputElement).value);

        // Module ModName
        const module_modname_dom  = module_dom.elements.namedItem("modulename") as HTMLInputElement
                                                                                    || throwf(new Error("WSC course get module, modname not found."));
        const module_out_modname =        module_modname_dom.value;

        // Module Intro/Description
        const module_description_dom = module_dom.elements.namedItem("introeditor[text]") as HTMLTextAreaElement
                                                                                    || throwf(new Error("WSC course get module, description not found."));
        if (module_in && module_in.mdl_course_module_instance && module_in.mdl_course_module_instance.intro != undefined) {
            module_description_dom.value = module_in.mdl_course_module_instance.intro;
        }
        const module_out_instance_intro = module_description_dom.value;

        // Module Name
        const module_name_dom = module_dom.elements.namedItem("name") as HTMLInputElement;
        // TODO: For label, instead of name field, use introeditor[text] field (without markup)?

        if (module_in && module_in.mdl_course_module_instance && module_in.mdl_course_module_instance.name != undefined) {
            module_name_dom.value = module_in.mdl_course_module_instance.name;
        }

        const module_out_instance_name = module_name_dom ? module_name_dom.value : "";


        // Module Completion
        const module_completion_dom = module_dom.elements.namedItem("completion") as HTMLInputElement;
        const module_completion: number = module_completion_dom ? parseInt(module_completion_dom.value) : 0;
        if (module_completion == 0 || module_completion == 1 || module_completion == 2) {  }
        else                                                                        { throw new Error("WSC course get module, completion value unexpected."); }
        // module_out_completion = module_completion;

        // For assignments
        const module_assignsubmission_file_enabled_x_dom = module_dom.elements.namedItem("assignsubmission_file_enabled") as HTMLInputElement;
        const module_assignsubmission_onlinetext_enabled_x_dom = module_dom.elements.namedItem("assignsubmission_onlinetext_enabled") as HTMLInputElement;

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
            id:         module_out_id,
            instance:   module_out_instance,
            course:     module_out_course,
            section:    module_out_section,
            mdl_modules_name: module_out_modname,
            mdl_course_module_instance: {
                name: module_out_instance_name,
                intro: module_out_instance_intro
            }
        };

        return {page: "mod-[a-z]+-mod", mdl_course_modules: module_out};
    }



    export type page_mod_feedback_edit_data = Page_Data_Base & {
        page: "mod-feedback-edit",
        location?: {pathname: "/mod/feedback/edit.php", search: {id: number, do_show: "edit"|"templates"}},
        mdl_course_modules?: { id?: number, mdl_course_module_instance?: { mdl_feedback_template_id?: number; } }
    };

    async function page_mod_feedback_edit(message: Page_Data_In_Base & DeepPartial<page_mod_feedback_edit_data>): Promise<page_mod_feedback_edit_data> {
       const template_id_dom = document.querySelector(":root #region-main form#mform2.mform select#id_templateid") as HTMLSelectElement;
       if (message && message.mdl_course_modules && message.mdl_course_modules.mdl_course_module_instance
            && message.mdl_course_modules.mdl_course_module_instance.hasOwnProperty("mdl_feedback_template_id")) {
            template_id_dom.value = "" + message.mdl_course_modules.mdl_course_module_instance.mdl_feedback_template_id;
            template_id_dom.dispatchEvent(new Event("change"));

       }
       return {page: "mod-feedback-edit"};
    }


    export type page_mod_feedback_use_templ_data = Page_Data_Base & {
        page: "mod-feedback-use_templ";
        location?: {pathname: "/mod/feedback/use_templ.php"}
        // mdl_course_modules: {x_submit: boolean;};
        // dom_submit: boolean
    };

    async function page_mod_feedback_use_templ(message: Page_Data_In_Base & DeepPartial<page_mod_feedback_use_templ_data>): Promise<page_mod_feedback_use_templ_data> {
       const submit_dom = document.querySelector(":root #region-main form#mform1.mform input#id_submitbutton") as HTMLInputElement;
       if (message && message.dom_submit) { // message.mdl_course_modules.x_submit) {
            submit_dom.click();
       }
       return {page: "mod-feedback-use_templ"};
    }







    async function page_onMessage(message: Page_Data_In, sender?: browser.runtime.MessageSender): Promise<Page_Data_Out> {
        if (sender && sender.tab !== undefined) { throw new Error("Unexpected message"); }
        return await page_get_set(message);
    }

    async function page_get_set(message_in: Page_Data_In): Promise<Page_Data_Out> {

        const message = message_in as Page_Data_In;



        let result: Page_Data;

        switch (window.document.body.id) {
            case "page-course-index":
            case "page-course-index-category":
                result = await page_course_index(message);
                break;
            case "page-course-view-onetopic":
            case "page-course-view-multitopic":
                result = await page_course_view(message);
                break;
            case "page-backup-backup":
                result = await page_backup_backup(message);
                break;
            case "page-backup-backupfilesedit":
                result = await page_backup_backupfilesedit(message);
                break;
            case "page-backup-restorefile":
                result = await page_backup_restorefile(message);
                break;
            case "page-backup-restore":
                result = await page_backup_restore(message);
                break;
            case "page-course-editsection":
                result = await page_course_editsection(message);
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
                result = await page_module_edit(message);
                break;
            case "page-mod-feedback-edit":
                result = await page_mod_feedback_edit(message);
                break;
            case "page-mod-feedback-use_templ":
                result = await page_mod_feedback_use_templ(message);
                break;
            default:
                result = {page: "*"};
                break;
        }

        (result as Page_Data_Out).moodle_page = {
                wwwroot:    window.location.origin,
                // location_pathname:  window.location.pathname,
                // location_search:    window.location.search,
                // location_hash:      window.location.hash,
                body_id:            window.document.body.getAttribute("id")             || throwf(new Error("WSC doc details get, body ID not found.")),
                body_class:         window.document.body.getAttribute("class")          || throwf(new Error("WSC doc details get, body class not found.")),
                sesskey:       (((window.document.querySelector(":root a.menu-action[data-title='logout,moodle']")
                                                                                        || throwf(new Error("WSC doc details get, couldn't get logout menu item.")) // Caught
                                    ) as HTMLAnchorElement
                                    ).search.match(/^\?sesskey=(\w+)$/)                || throwf(new Error("WSC doc details get, session key not found."))
                                    )[1],
                // error_message:      error_message_dom ? error_message_dom.textContent || "" : undefined,

        };


        return result as Page_Data_Out;
    }



    export async function page_init(): Promise<void>/*{status: boolean}*/ {
        browser.runtime.onMessage.addListener(page_onMessage);
        // return {status: true};
        // return c_on_call({});
        let message: Page_Data_Out|Errorlike;
        try {
            message = await page_onMessage({});
        } catch (e) {
            message = {name: "Error", message: e.message, fileName: e.fileName, lineNumber: e.lineNumber};
        }
        void browser.runtime.sendMessage(message);
    }



}


void MJS.page_init();
