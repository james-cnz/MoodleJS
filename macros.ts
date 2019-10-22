/**
 * Moodle JS Macro Scripts
 */


namespace MJS {




    export class TabData {


        public static escapeHTML(text: string) {
            return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");  // TODO: Line breaks?  nbsp?
        }


        private static json_to_search_params(search_json: {[index: string]: number|string}): string {
            const search_obj: URLSearchParams = new URLSearchParams();
            for (const key in search_json)
                if (search_json.hasOwnProperty(key) && search_json[key] != undefined) {
                    search_obj.append(key, "" + search_json[key]);
            }
            return search_obj.toString();
        }

        private static compare_skel(data: any, skel: any): boolean {
            if (skel == undefined) { return true; }
            else if (typeof(skel) == "function")  {
                throw new Error("Can't compare functions");
            } else if (typeof(skel) == "object") {
                for (const prop in skel) if (skel.hasOwnProperty(prop) && skel.prop != undefined) {
                    if (!data.hasOwnProperty(prop) || !this.compare_skel(data[prop], skel[prop])) { return false; }
                }
                return true;
            } else {
                return data == skel;
            }
        }



        public page_wwwroot:       string;
        public page_sesskey:       string;

        public macro_state:        number = 0;     // -1 error / 0 idle / 1 running / 2 running & awaiting load
        // macro_callstack:    string[3]       // 0: macro  1: macro step  2: tabdata function
        public macro_error:        Errorlike|null;
        public macro_progress:     number = 100;
        public macro_progress_max: number = 100;
        public macro_cancel:       boolean = false;

        public macros: {[index: string] : Macro} = {
            new_course: new New_Course_Macro(this),
            index_rebuild: new Index_Rebuild_Macro(this),
            new_section: new New_Section_Macro(this),
            new_topic: new New_Topic_Macro(this),
            test: new Test_Macro(this)
        };

        private page_details:      Page_Data_Out;
        private page_tab_id:        number;

        private page_load_wait:     number = 0;
        private page_message:       Page_Data_Out|Errorlike|null = null;
        private page_is_loaded:     boolean = false;

        private popup:              Popup;


        constructor(tab_id: number) {
            this.page_tab_id = tab_id;
            void this.init();
        }


        public async init() {
            this.page_load_wait = 0;
            this.page_message = null;
            this.page_is_loaded = true;
            this.macro_error = null;
            this.macro_progress = 100;
            this.macro_progress_max = 100;
            this.macro_cancel = false;
            try {
                this.macro_state = 1;
                this.page_details = await this.page_call({});
                this.page_wwwroot = this.page_details.page_window.location_origin;
                this.page_sesskey = this.page_details.page_window.sesskey;
                this.macro_state = 0;
            } catch (e) {
                this.macro_state = 0;
                this.page_details = null;
                this.page_wwwroot = null;
                this.page_sesskey = null;
            }
            this.macros_init(this.page_details);
            this.macro_state = 0;
            this.update_ui();
            // console.log("init finished.");
        }


        public update_ui(): void {
            /*
            if (this.macro_state == 0) {
                browser.browserAction.setBadgeBackgroundColor({color: "green", tabId: this.page_tab_id});
                browser.browserAction.setBadgeText({text: "", tabId: this.page_tab_id});
            } else if (this.macro_state > 0) {
                browser.browserAction.setBadgeBackgroundColor({color: "green", tabId: this.page_tab_id});
                browser.browserAction.setBadgeText({text: ">", tabId: this.page_tab_id});
            } else {
                browser.browserAction.setBadgeBackgroundColor({color: "red", tabId: this.page_tab_id});
                browser.browserAction.setBadgeText({text: "X", tabId: this.page_tab_id});
            }
            */
            try {
                this.popup.update();
            } catch (e) { }
        }


        public async page_call<T extends Page_Data>(message: DeepPartial<T> & Page_Data_In_Base): Promise<T & Page_Data_Out_Base> {

            (this.macro_state == 1)                                             || throwf(new Error("Page call:\nUnexpected state."));

            if (message.dom_submit) {
                this.macro_state = 2;
                this.page_is_loaded = false;
                this.page_message = null;
            }
            const result = await browser.tabs.sendMessage(this.page_tab_id, message) as (T & Page_Data_Out_Base)|Errorlike;
            if (is_Errorlike(result))                                            { throw (new Error((result as Errorlike).message)); }

            return result;
        }


        public async page_load<T extends Page_Data>(pathname: string, search: {[index: string]: number|string},
                                page_data: DeepPartial<T>,
                                count: number = 1): Promise<T & Page_Data_Out_Base> {

            (this.macro_state == 1)                                             || throwf(new Error("Page load:\nUnexpected state."));
            (pathname.match(/(?:\/[a-z]+)+\.php/))                              || throwf(new Error("Page load:\nPathname unexpected."));

            this.macro_state = 2;
            this.page_is_loaded = false;
            this.page_message = null;

            await browser.tabs.update(this.page_tab_id, {url: this.page_wwwroot + pathname + "?" + TabData.json_to_search_params(search)});
            return await this.page_loaded(page_data, count);
        }


        public async page_loaded<T extends Page_Data>(page_data: DeepPartial<T>,
            count: number = 1): Promise<T & Page_Data_Out_Base> {

            (this.macro_state == 2)                                             || throwf(new Error("Page loaded:\nUnexpected state."));

            this.page_load_wait  = 0;
            do {
                // (this.page_load_wait < 600)                                     || throwf(new Error("MJS page loaded: Timed out."));
                await sleep(100);
                if (this.macro_cancel)                                          { throw new Error("Cancelled"); }
                this.page_load_wait += 1;
                if (this.page_load_wait <= count * 10) {  // Assume a step takes 1 second
                    this.page_load_count(1 / 10);
                }
            } while (!(this.page_is_loaded && this.page_message) && !(this.page_message && is_Errorlike(this.page_message)));
            this.macro_state = 1;
            if (is_Errorlike(this.page_message))                                 { throw new Error(this.page_message.message, this.page_message.fileName, this.page_message.lineNumber); }
            this.page_details = this.page_message;
            this.page_message = null;
            if (this.page_load_wait <= count * 10) {
                this.page_load_count(count - this.page_load_wait / 10);
            }
            if (!this.page_load_match<T>(this.page_details, page_data))                   { throw new Error("Page loaded:\nBody ID or class unexpected."); }
            this.macro_state = 1;
            return this.page_details;
        }


        public page_load_count(count: number = 1): void {
            this.macro_progress += count;
            try {
                this.popup.update_progress();
            } catch (e) { }
        }


        public onMessage(message: Page_Data_Out|Errorlike, _sender: browser.runtime.MessageSender) {

            if (!this.page_message || this.macro_state == 0) {
                this.page_message = message;
                if (is_Errorlike(message)) {
                    this.page_wwwroot = null;
                    this.page_sesskey = null;
                } else {
                    this.page_details = message;
                    // console.log("updated page details");
                    this.page_wwwroot = this.page_details.page_window.location_origin;
                    this.page_sesskey = this.page_details.page_window.sesskey;
                }

                if (this.page_is_loaded ) {

                    // console.log("*** late page message ***");
                    if (this.macro_state == 0) { this.macros_init(this.page_details); }
                    this.update_ui();
                }
            } else if (!is_Errorlike(this.page_message)) {
                this.page_message = new Error("Unexpected page message");
            }

        }


        public onTabUpdated(_tab_id: number, update_info: Partial<browser.tabs.Tab>, _tab: browser.tabs.Tab): void {

            if (update_info && update_info.status) {

                if (!this.page_is_loaded || this.macro_state == 0) {

                    if (update_info.status == "loading") {
                        if (this.macro_state == 0) {
                            this.page_is_loaded = false;
                            this.page_message = null;
                        }
                    } else if (update_info.status == "complete") {
                        this.page_is_loaded = true;

                        if (!this.page_message) {
                            if (this.macro_state == 0) { this.page_details = null; this.macros_init(this.page_details); }
                            else if (this.macro_state == 2) {
                                // console.log("*** missing page message ***");
                            }
                        } else {
                            if (this.macro_state == 0) { this.macros_init(this.page_details); }
                            this.update_ui();
                        }
                    }
                } else if (!this.page_message || !is_Errorlike(this.page_message)) {
                    // this.m_state = -1;
                    this.page_message = new Error("Unexpected tab update");

                }

                this.update_ui();

            }

        }


        private macros_init(page_details: Page_Data_Out) {
            this.macros.new_course.init(page_details);
            this.macros.index_rebuild.init(page_details);
            this.macros.new_section.init(page_details);
            this.macros.new_topic.init(page_details);
        }


        private page_load_match<T extends Page_Data>(page_details: any, page_data: DeepPartial<T>):
            page_details is T & Page_Data_Out_Base {
            let result = true;
            if (!page_data.page || page_details.page_window.body_id.match(RegExp("^page-" + page_data.page + "$"))) { /* OK */ } else    { result = false; }
            /*for (const prop in body_class) if (body_class.hasOwnProperty(prop)) {
                if ((" " + this.page_details.page_window.body_class + " ").match(" " + prop + (body_class[prop] ? ("-" + body_class[prop]) : "") + " "))
                    {  OK  }
                else
                    { result = false; }
            }*/
            return result;
        }


    }




    abstract class Macro {


        public prereq: boolean = false;


        protected tabdata: TabData;

        protected progress_max: number;

        protected page_details: Page_Data_Out;

        constructor(new_tabdata: TabData) {
            this.tabdata = new_tabdata;
            this.prereq = false;
        }


        public abstract init(page_details: Page_Data_Out): void;


        public async run(): Promise<void> {

            if (this.tabdata.macro_state != 0) {
                return;
            }

            // this.init();

            this.tabdata.macro_cancel = false;
            this.tabdata.macro_state = 1;
            this.tabdata.macro_progress = 0;
            this.tabdata.macro_progress_max =  this.progress_max;

            try {
                await this.content();
            } catch (e) {
                if (e.message != "Cancelled") {
                    this.tabdata.macro_state = -1;
                    this.tabdata.macro_error = e;
                    this.tabdata.update_ui();
                    return;
                }
            }

            this.tabdata.macro_state = 0;
            this.tabdata.macro_progress = this.tabdata.macro_progress_max;
            this.tabdata.macros_init(this.tabdata.page_details);
            try {
                this.tabdata.popup.close();
            } catch (e) {
                // Do nothing
            }

        }




        /*
        protected get page_details(): Page_Data_Out {
            return this.tabdata.page_details;
        }


        protected set page_details(page_details_in: Page_Data_Out) {
            (page_details_in == this.tabdata.page_details)                      || throwf(new Error("Page details:\nPage details mismatch."));
        }
        */


       public async page_call<T extends Page_Data>(message: DeepPartial<T> & Page_Data_In_Base): Promise<T & Page_Data_Out_Base> {
            return await this.tabdata.page_call(message);
        }


        public async page_load<T extends Page_Data>(pathname: string, search: {[index: string]: number|string},
            page_data: DeepPartial<T>,
            count: number = 1): Promise<T & Page_Data_Out_Base> {
            return await this.tabdata.page_load(pathname, search, page_data, count);
        }


        public async page_loaded<T extends Page_Data>(page_data: DeepPartial<T>,
            count: number = 1): Promise<T & Page_Data_Out_Base> {
            return await this.tabdata.page_loaded(page_data, count);
        }


        protected abstract async content(): Promise<void>;


    }




    export class New_Course_Macro extends Macro {


        // public prereq:             boolean;

        public new_course: DeepPartial<MDL_Course> & {
            fullname: string;
            shortname: string;
            startdate: number;
        };

        private course_template_id: number;
        private category_id:        number;
        private category_name:      string;


        public init(page_details: Page_Data_Out) {

            // console.log("new course pre starting");
            this.prereq = false;

            this.page_details = page_details;

            if (!this.page_details || !this.page_details.hasOwnProperty("page") || this.page_details.page != "course-index(-category)?")     { return; }

            if (this.page_details.page_window.body_id != "page-course-index-category") { return; }

            this.category_id  = this.page_details.mdl_course_categories.id;
            this.category_name = this.page_details.mdl_course_categories.name;

            if (this.tabdata.page_wwwroot == "https://otagopoly-moodle.testing.catlearn.nz" ) {
                this.course_template_id = 6548;
            } else if (this.tabdata.page_wwwroot == "https://moodle.op.ac.nz") {
                this.course_template_id = 6548;
            } else                                                                      { return; }

            this.progress_max = 17 + 1;
            this.prereq = true;

        }


        protected async content() {  // TODO: Set properties.

            const name = this.new_course.fullname;
            const shortname = this.new_course.shortname;
            const startdate = this.new_course.startdate;

            // Get template course context (1 load)
            this.page_details = await this.page_load("/course/view.php", {id: this.course_template_id, section: 0},
                {page: "course-view-[a-z]+", mdl_course: {id: this.course_template_id}},
            );
            const source_context_match = this.page_details.page_window.body_class.match(/(?:^|\s)context-(\d+)(?:\s|$)/)
                                                                                        || throwf(new Error("New course macro, get template:\nContext not found."));
            const source_context = parseInt(source_context_match[1]);

            // Load course restore page (1 load)
            this.page_details = await this.page_load(
                "/backup/restorefile.php", {contextid: source_context},
                {page: "backup-restorefile", mdl_course: {id: this.course_template_id}},
            );

            // Click restore backup file (1 load)
            this.page_details = await this.page_call<page_backup_restorefile_data>({page: "backup-restorefile", dom_submit: "restore"});
            this.page_details = await this.page_loaded<page_backup_restore_data>({page: "backup-restore", stage: 2, mdl_course: {template_id: this.course_template_id}});

            // Confirm (1 load)
            (this.page_details.stage == 2)                                      || throwf(new Error("New course macro, confirm:\nStage unexpected."));
            this.page_details = await this.page_call({page: "backup-restore", stage: 2, dom_submit: "stage 2 submit"});
            this.page_details = await this.page_loaded<page_backup_restore_data>({page: "backup-restore", mdl_course: {template_id: this.course_template_id}});

            // Destination: Search for category (1 load)
            (this.page_details.stage == 4)                                      || throwf(new Error("New course macro, destination:\nStage unexpected."));
            this.page_details = await this.page_call({page: "backup-restore", stage: 4, mdl_course_categories: {name: this.category_name}, dom_submit: "stage 4 new cat search"});
            this.page_details = await this.page_loaded({page: "backup-restore"});  // TODO: Add details

            // Destination: Select category (1 load)
            this.page_details = await this.page_call({page: "backup-restore", stage: 4, mdl_course_categories: {id: this.category_id}, dom_submit: "stage 4 new continue"});
            this.page_details = await this.page_loaded<page_backup_restore_data>({page: "backup-restore", stage: 4, mdl_course: {template_id: this.course_template_id}});

            // Restore settings (1 load)
            if (this.page_details.stage != 4)                                    { throw new Error("New course macro, restore settings:\nStage unexpected."); }
            this.page_details = await this.page_call({page: "backup-restore", stage: 4, restore_settings: {users: false}});
            this.page_details = await this.page_call({page: "backup-restore", dom_submit: "stage 4 settings submit"});
            this.page_details = await this.page_loaded<page_backup_restore_data>({page: "backup-restore", mdl_course: {template_id: this.course_template_id}});

            // Course settings (1 load)
            (this.page_details.stage == 8)                                      || throwf(new Error("New course macro, course settings:\nStage unexpected."));
            const course: DeepPartial<MDL_Course> = {fullname: name, shortname: shortname, startdate: startdate};
            this.page_details = await this.page_call({page: "backup-restore", stage: 8, mdl_course: course, dom_submit: "stage 8 submit"});
            this.page_details = await this.page_loaded<page_backup_restore_data>({page: "backup-restore", mdl_course: {template_id: this.course_template_id}});

            // Review & Process (~5 loads)
            (this.page_details.stage == 16)                                     || throwf(new Error("New course macro, review & process:\nStage unexpected"));
            this.page_details = await this.page_call({page: "backup-restore", stage: 16, dom_submit: "stage 16 submit"});
            this.page_details = await this.page_loaded<page_backup_restore_data>({page: "backup-restore", mdl_course: {template_id: this.course_template_id}}, 5);

            // Complete--Go to new course (1 load)
            (this.page_details.stage == null)                                   || throwf(new Error("New course macro, complete:\nStage unexpected."));
            const course_id = (this.page_details.mdl_course as DeepPartial<MDL_Course>).id as number;
            this.page_details = await this.page_call({page: "backup-restore", stage: null, dom_submit: "stage complete submit"});
            this.page_details = await this.page_loaded<page_course_view_data>({page: "course-view-[a-z]+", mdl_course: {id: course_id}});

            // TODO: Turn editing on (1 load).
            if (!this.page_details.page_window.body_class.match(/\bediting\b/)) {
                this.page_details = await this.page_load<page_course_view_data>(
                    "/course/view.php", {id: course_id, section: 0, sesskey: this.tabdata.page_sesskey, edit: "on"},
                    {page: "course-view-[a-z]+", mdl_course: {id: course_id}},
                );
            } else {
                this.tabdata.page_load_count(1);
            }
            (this.page_details.page_window.body_class.match(/\bediting\b/))     || throwf(new Error("New course macro, turn editing on:\nEditing not on."));

            // Fill in course name (2 loads)
            const section_0_id = this.page_details.mdl_course_sections.id       || throwf(new Error("New course macro, fill in course name:\nID not found."));
            this.page_details = await this.page_load<page_course_editsection_data>("/course/editsection.php", {id: section_0_id}, // TODO: Needs editing on.
                {page: "course-editsection", mdl_course: {id: course_id}},
            );
            // TODO: Crashes around here with "can't access dead object"?
            const desc = this.page_details.mdl_course_sections.summary.replace(/\[Course Name\]/g, this.new_course.fullname);
            this.page_details = await this.page_call({page: "course-editsection", mdl_course_sections: {summary: desc}, dom_submit: true});
            this.page_details = await this.page_loaded({page: "course-view-[a-z]+", mdl_course: {id: course_id}});

        }


    }




    export class Index_Rebuild_Macro extends Macro {


        // prereq: boolean;

        private course_id: number;
        private modules_tab_num: number;
        private last_module_tab_num: number;


        public init(page_details: Page_Data_Out) {

            this.prereq = false;

            this.page_details = page_details;

            // Check course type
            if (!this.page_details || this.page_details.page != "course-view-[a-z]+")                      { return; }
            const course = this.page_details.mdl_course;
            if (!course || course.format != "onetopic" || !course.id) {  return; }
            this.course_id = course.id;

            // Check editing on
            if (!this.page_details.page_window || !this.page_details.page_window.body_class || !this.page_details.page_window.body_class.match(/\bediting\b/)) {
                // console.log("index rebuild pre: editing not on");
                return;
            }

            // Find Modules tab number
            const course_contents = course.mdl_course_sections;
            let modules_tab_num: number|undefined|null = null;
            let last_module_tab_num: number|undefined|null = null;
            for (const section of course_contents) {
                if ((section.x_options.level <= 0) && (section.section as number <= this.page_details.mdl_course_sections.section) && (section.name.toUpperCase().trim() == "MODULES")) {
                    modules_tab_num = section.section;
                    last_module_tab_num = modules_tab_num;
                } else if (last_module_tab_num && section.x_options.level > 0) { // TODO: Need to scrape level property.
                    last_module_tab_num = section.section;
                }
            }
            if (modules_tab_num) {  } else                                        {  return; }
            if (this.page_details.mdl_course_sections.section <= last_module_tab_num)
            { } else { return; }
            this.modules_tab_num = modules_tab_num;
            this.last_module_tab_num = last_module_tab_num;

            this.progress_max = this.last_module_tab_num - this.modules_tab_num + 3 + 1;
            this.prereq = true;

        }


        protected async content() {

            // TODO: Don't include hidden tabs or topic headings?

            const parser = new DOMParser();

            // Get list of sections (1 load)
            this.page_details = await this.page_load<page_course_view_data>("/course/view.php", {id: this.course_id, section: this.modules_tab_num},
                {page: "course-view-[a-z]+", mdl_course: {id: this.course_id}},
            );
            const modules_list = this.page_details.mdl_course_sections.mdl_course_modules as MDL_Course_Modules[];
            (modules_list.length == 1)                                          || throwf(new Error("Index rebuild macro, get list:\nExpected exactly one resource in Modules tab."));
            const modules_index = modules_list[0];


            // Get section contents (1 load per section)
            let index_html = '<div class="textblock">\n';
            // modules_list.shift();
            for (let section_num = this.modules_tab_num + 1; section_num <= this.last_module_tab_num; section_num++) {
                // const section_num = section.section                                     || throwf(new Error("Module number not found."));
                this.page_details = await this.page_load<page_course_view_data>("/course/view.php", {id: this.course_id, section: section_num},
                                    {page: "course-view-[a-z]+", mdl_course: {id: this.course_id}});
                const section_full = this.page_details.mdl_course_sections;
                const section_name = (parser.parseFromString(section_full.summary as string, "text/html").querySelector(".header1")
                                                                                        || throwf(new Error("Index rebuild macro, get section:\nModule name not found."))
                                    ).textContent                                      || throwf(new Error("Index rebuild macro, get section:\nModule name content not found."));
                index_html = index_html
                            + '<a href="' + this.tabdata.page_wwwroot + "/course/view.php?id=" + this.course_id + "&section=" + section_num + '"><b>' + TabData.escapeHTML(section_name.trim()) + "</b></a>\n"
                            + "<ul>\n";
                for (const mod of section_full.mdl_course_modules as DeepPartial<MDL_Course_Modules>[]) {
                    // parse description
                    const mod_desc = parser.parseFromString((mod.mdl_course_module_instance as DeepPartial<MDL_Course_Module_Instance>).intro || "", "text/html");
                    const part_name = mod_desc.querySelector(".header2, .header2gradient");
                    if (part_name) {
                        index_html = index_html
                                    + "<li>"
                                    + TabData.escapeHTML((part_name.textContent                 || throwf(new Error("Index rebuild macro, get section:\nCouldn't get text content."))
                                    ).trim())
                                    + "</li>\n";
                    }
                }
                index_html = index_html
                            + "</ul>\n"
                            + "<br />\n";
            }
            index_html = index_html
                        + "</div>\n";

            // Update TOC (2 loads)
            this.page_details = await this.page_load("/course/mod.php", {sesskey: this.tabdata.page_sesskey, update: modules_index.id},
                                    {page: "mod-[a-z]+-mod", mdl_course: {id: this.course_id}});
            this.page_details = await this.page_call({page: "mod-[a-z]+-mod", mdl_course_modules: {course: this.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
                mdl_course_module_instance: {intro: index_html}}, dom_submit: true});
            this.page_details = await this.page_loaded({page: "course-view-[a-z]+", mdl_course: {id: this.course_id}});

        }


    }




    export class New_Section_Macro extends Macro {


        // prereq:                 boolean;


        public new_section: DeepPartial<MDL_Course_Sections> & {
            fullname: string;
            name: string;
        };


        private feedback_template_id:   number;
        private course_id:              number;
        private new_section_pos:        number|null;


        public init(page_details: Page_Data_Out) {

            // console.log("new section pre starting");
            this.prereq = false;
            // Get site details
            // TODO: Also customise image link per site

            this.page_details = page_details;

            // Check page type
            if (!this.page_details || this.page_details.page != "course-view-[a-z]+") {
                // console.log("new section pre: wrong page type");
                return;
            }

            // let feedback_template_id: number;
            if (this.tabdata.page_wwwroot == "https://otagopoly-moodle.testing.catlearn.nz" ) {
                this.feedback_template_id = 59;
            } else if (this.tabdata.page_wwwroot == "https://moodle.op.ac.nz") {
                this.feedback_template_id = 59;
            } else                                                                      { return; }

            // Check editing on
            if (!this.page_details.page_window || !this.page_details.page_window.body_class || !this.page_details.page_window.body_class.match(/\bediting\b/)) {
                // console.log("new section pre: editing not on");
                return;
            }

            // Get course details
            // console.log("get course details");
            const course = (this.page_details as page_course_view_data).mdl_course; // (await this.page_call({})).mdl_course;
            if (!course) { /*console.log("new section pre: couldn't get course details");*/ return; }
            this.course_id = course.id;

            if (course.format == "onetopic") {  } else                            { return; }

            // Find Modules tab number
            const course_contents = course.mdl_course_sections;
            let modules_tab_num: number|undefined|null = null;
            let last_module_tab_num: number|undefined|null = null;
            for (const section of course_contents) {
                if (section.x_options.level <= 0 && section.section <= this.page_details.mdl_course_sections.section && section.name.toUpperCase().trim() == "MODULES") {
                    modules_tab_num = section.section;
                    last_module_tab_num = modules_tab_num;
                } else if (last_module_tab_num && section.x_options.level > 0) { // TODO: Need to scrape level property.
                    last_module_tab_num = section.section;
                }
            }
            if (modules_tab_num) {  } else                                        { return; }
            if (this.page_details.mdl_course_sections.section <= last_module_tab_num)
            { } else { return; }
            this.new_section_pos = last_module_tab_num + 1;

            this.progress_max = 17 + 1;
            this.prereq = true;
            // console.log("new section pre success");
        }


        protected async content() {

            const name = this.new_section.fullname;
            const short_name = this.new_section.name;

            // Add new tab (1 load)
            this.page_details = await this.page_load<page_course_view_data>(  // TODO: Fix for flexsections?
                "/course/changenumsections.php", {courseid: this.course_id, increase: 1, sesskey: this.tabdata.page_sesskey, insertsection: 0},
                {page: "course-view-[a-z]+", mdl_course: {id: this.course_id}},
            );
            let new_section = this.page_details.mdl_course_sections;

            // Move new tab (1 load)
            // console.log("Move new tab");
            this.page_details = await this.page_load(
                "/course/view.php", {id: this.course_id, section: new_section.section, sesskey: this.tabdata.page_sesskey, move: this.new_section_pos - new_section.section
                                                                                    /*|| throwf(new Error("WS course section edit, no amount specified."))*/},
                {page: "course-view-[a-z]+", mdl_course: {id: this.course_id}},
            );
            new_section.section = this.new_section_pos;

            // Set new tab details (2 loads)
            this.page_details = await this.page_load(
                "/course/editsection.php", {id: new_section.id, sr: new_section.section,
                                                                                    /*|| throwf(new Error("WS course section edit, no amount specified."))*/},
                                                                                    {page: "course-editsection", mdl_course: {id: this.course_id}},
                                                                                    );
            this.page_details = await this.page_call({page: "course-editsection", mdl_course_sections: {id: new_section.id, name: short_name, x_options: {level: 1},
                summary:
                `<div class="header1"> <i class="fa fa-list" aria-hidden="true"></i> ${name}</div>

                <p></p>

                <img src="https://moodle.op.ac.nz/pluginfile.php/812157/course/section/106767/nature-sky-clouds-flowers.jpg" alt="Generic sky" style="float: right; margin-left: 5px; margin-right: 5px;" width="240" height="180" class="img-responsive" />

                <p>[Intro to the module goes here.
                Lorem ipsum dolor sit amet, consectetur adipiscing elit.
                Cras iaculis mollis efficitur.
                Praesent ipsum diam, dignissim et orci et, tempor fringilla lectus.
                Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia Curae; Proin sed quam pharetra, gravida odio iaculis, fermentum turpis.
                Etiam vel tincidunt justo, at fringilla sem.]</p>

                <p>This module will provide you with information, learning activities, and resources that support your classroom and other aspects (e.g. projects, work experiences) of the course work.</p>

                <p>If this is your first visit, we suggest that you work through each topic in the sequence set out below, starting with <strong>[xxxxxxxxx]</strong>.
                As you work through the topics, please access the learning activities below, as these are an essential part of your learning in this programme.</p>

                <p>We recommend that you visit this module on a regular basis, to complete the activities and to self-test your increasing knowledge and skills.</p>`.replace(/^        /gm, ""),
            }, dom_submit: true});
            this.page_details = await this.page_loaded({page: "course-view-[a-z]+", mdl_course: {id: this.course_id}});

            // Add pre-topic message (2 loads)
            // Body ID or class unexpected.
            this.page_details = await this.page_load("/course/mod.php", {id: this.course_id, sesskey: this.tabdata.page_sesskey, sr: new_section.section, add: "label", section: new_section.section},
                                    {page: "mod-[a-z]+-mod", mdl_course: {id: this.course_id}},
                                    );
            this.page_details = await this.page_call({page: "mod-[a-z]+-mod", mdl_course_modules: {course: this.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
                mdl_course_module_instance: {intro:
                `<p></p>

                <p>After you have worked through all of the above topics, and your facilitator provides you with further information in class,
                you're now ready to demonstrate evidence of what you have learnt in this module.
                Please click on the <strong>Assessments</strong> tab above for further information.</p>`.replace(/^        /gm, "")},
                }, dom_submit: true});
            this.page_details = await this.page_loaded({page: "course-view-[a-z]+", mdl_course: {id: this.course_id}});

            // Add blank line (2 loads)
            this.page_details = await this.page_load("/course/mod.php", {id: this.course_id, sesskey: this.tabdata.page_sesskey, sr: new_section.section, add: "label", section: new_section.section},
            {page: "mod-[a-z]+-mod", mdl_course: {id: this.course_id}},
            );
            this.page_details = await this.page_call({page: "mod-[a-z]+-mod", mdl_course_modules: {course: this.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
            mdl_course_module_instance: {intro: ""}, }, dom_submit: true});
            this.page_details = await this.page_loaded({page: "course-view-[a-z]+", mdl_course: {id: this.course_id}});

            // Add feedback topic (2 loads)
            this.page_details = await this.page_load("/course/mod.php", {id: this.course_id, sesskey: this.tabdata.page_sesskey, sr: new_section.section, add: "label", section: new_section.section},
                {page: "mod-[a-z]+-mod", mdl_course: {id: this.course_id}},
            );
            this.page_details = await this.page_call({page: "mod-[a-z]+-mod", mdl_course_modules: {course: this.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
            mdl_course_module_instance: {intro:
                `<p></p>

                <p><strong>YOUR FEEDBACK</strong></p>

                <p>We appreciate your feedback about your experience with working through this module.
                Please click the 'Your feedback' link below if you wish to respond to a five-question survey.
                Thanks!</p>`.replace(/^        /gm, "")},
            }, dom_submit: true});
            this.page_details = await this.page_loaded({page: "course-view-[a-z]+", mdl_course: {id: this.course_id}});

            // Add feedback activity (2 load)
            this.page_details = await this.page_load("/course/mod.php", {id: this.course_id, sesskey: this.tabdata.page_sesskey, sr: new_section.section, add: "feedback", section: new_section.section},
                {page: "mod-[a-z]+-mod", mdl_course: {id: this.course_id}},
            );
            this.page_details = await this.page_call({page: "mod-[a-z]+-mod", mdl_course_modules: {course: this.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
            mdl_course_module_instance: {name: "Your feedback", intro:
                `<div class="header2"> <i class="fa fa-bullhorn" aria-hidden="true"></i> FEEDBACK</div>

                <div class="textblock">

                <p><strong>DESCRIPTION</strong></p>

                <p>Please help us improve this learning module by answering five questions about your experience.
                This survey is anonymous.</p>
                </div>`.replace(/^        /gm, "")}, }, dom_submit: true});
            this.page_details = await this.page_loaded({page: "course-view-[a-z]+", mdl_course: {id: this.course_id}});
            new_section = (this.page_details as page_course_view_data).mdl_course_sections;
            let feedback_act: DeepPartial<MDL_Course_Modules>|null = null;
                for (const module of new_section.mdl_course_modules) {
                    if (!feedback_act || module.id > feedback_act.id) {
                        feedback_act = module;
                    }
                }

            // Configure Feedback activity (3 loads?)
            this.page_details = await this.page_load("/mod/feedback/edit.php", {id: feedback_act.id, do_show: "templates"},
                                {page: "mod-feedback-edit", mdl_course_modules: {id: feedback_act.id}});
            this.page_details = await this.page_call({page: "mod-feedback-edit", mdl_course_modules: { mdl_course_module_instance: {mdl_feedback_template_id: this.feedback_template_id}}, dom_submit: true});  // TODO: fix;
            this.page_details = await this.page_loaded({page: "mod-feedback-use_templ"});
            this.page_details = await this.page_call({page: "mod-feedback-use_templ", dom_submit: true});
            this.page_details = await this.page_loaded({page: "mod-feedback-edit"});

            // Add footer (2 loads).
            this.page_details = await this.page_load("/course/mod.php", {id: this.course_id, sesskey: this.tabdata.page_sesskey, sr: new_section.section, add: "label", section: new_section.section},
            {page: "mod-[a-z]+-mod", mdl_course: {id: this.course_id}},
            );
            this.page_details = await this.page_call({page: "mod-[a-z]+-mod", mdl_course_modules: {course: this.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
            mdl_course_module_instance: {intro:
                `<p></p>

                <p><span style="font-size: xx-small;">
                Image: <a href="https://stock.tookapic.com/photos/12801" target="_blank">Blooming</a>
                by <a href="https://stock.tookapic.com/pawelkadysz" target="_blank">Paweł Kadysz</a>,
                licensed under <a href="https://creativecommons.org/publicdomain/zero/1.0/deed.en" target="_blank">CC0</a>
                </span></p>`.replace(/^        /gm, "")},
            }, dom_submit: true});
            this.page_details = await this.page_loaded({page: "course-view-[a-z]+", mdl_course: {id: this.course_id}});

        }


    }




    export class New_Topic_Macro extends Macro {


        // prereq:         boolean;

        public new_topic_name: string;

        private course_id:      number;
        private section_num:    number;
        private mod_move_to:    number;
        private topic_first:    boolean;


        public init(page_details: Page_Data_Out) {

            // console.log("new topic pre starting");
            this.prereq = false;

            this.page_details = page_details;

            // var doc_details = ws_page_call({wsfunction: "x_doc_get_details"});
            if (!this.page_details || this.page_details.page != "course-view-[a-z]+")                      { return; }
            const course = this.page_details.mdl_course;
            if (course && course.hasOwnProperty("format") && course.format == "onetopic" && course.id) {  } else { return; }
            this.course_id = course.id;


            // Check editing on
            if (!this.page_details.page_window || !this.page_details.page_window.body_class || !this.page_details.page_window.body_class.match(/\bediting\b/)) {
                // console.log("new topic pre: editing not on");
                return;
            }

            // const section_url = await this.page_call({id_act: "* get_element_attribute", selector: "#page-navbar a[href*='section=']", attribute: "href"})
            //                                                                            || throwf(new Error("Section breadcrumb not found."));
            // const section_match = section_url.match(/^(https?:\/\/[a-z\-.]+)\/course\/view.php\?id=(\d+)&section=(\d+)$/)
            //                                                                            || throwf(new Error("Section number not found."));
            // const section_num = parseInt(section_match[3]);

            // let section = (await ws_call({wsfunction: "core_course_get_contents", courseid: course.id, options: [{name: "sectionnumber", value: section_num}]}))[0];
            const section = this.page_details.mdl_course_sections;
            // const section_num = section.section;
            this.section_num = section.section;

            let mod_pos = section.mdl_course_modules.length - 1;
            let mod_match_pos = 3;
            // let match_ok = true;

            while (mod_pos > -1 && mod_match_pos > -1) { // } && match_ok) {
                if (mod_match_pos == 3 && section.mdl_course_modules[mod_pos].mdl_modules_name == "label" && section.mdl_course_modules[mod_pos].mdl_course_module_instance.name.toUpperCase().match(/\bIMAGE\b/)) {
                    mod_match_pos -= 1;
                    mod_pos -= 1;
                } else if (mod_match_pos == 3) {
                    mod_match_pos -= 1;
                } else if (mod_match_pos == 2 && section.mdl_course_modules[mod_pos].mdl_modules_name == "feedback" && section.mdl_course_modules[mod_pos].mdl_course_module_instance.name.toUpperCase().match(/\bFEEDBACK\b/)) {
                    mod_match_pos -= 1;
                    mod_pos -= 1;
                } else if (mod_match_pos == 1 && section.mdl_course_modules[mod_pos].mdl_modules_name == "label" && section.mdl_course_modules[mod_pos].mdl_course_module_instance.name.toUpperCase().match(/\bFEEDBACK\b/)) {
                    mod_match_pos -= 1;
                    mod_pos -= 1;
                } else if (mod_match_pos == 0 && section.mdl_course_modules[mod_pos].mdl_modules_name == "label" && section.mdl_course_modules[mod_pos].mdl_course_module_instance.name == "") {
                    mod_pos -= 1;
                } else if (mod_match_pos == 0 && section.mdl_course_modules[mod_pos].mdl_modules_name == "label" && section.mdl_course_modules[mod_pos].mdl_course_module_instance.name.replace(/\s+/g, " ").toUpperCase().match(/\bASSESSMENTS TAB\b/)) {
                    mod_match_pos -= 1;
                    mod_pos -= 1;
                } else {
                    break;
                    // match_ok = false;
                }
            }

            if (mod_match_pos < 0) {  } else                                      { return; }

            this.mod_move_to = section.mdl_course_modules[mod_pos + 1].id;

            this.topic_first = (mod_pos < 0) ? true : false;

            this.progress_max = 12 + 1;
            this.prereq = true;
            // console.log("new topic pre success");
        }


        protected async content() {

            const name = this.new_topic_name;

            // Create topic heading (4 loads?)
            this.page_details = await this.page_load("/course/mod.php", {id: this.course_id, sesskey: this.tabdata.page_sesskey, sr: 0 /* TODO: remove? */, add: "label", section: this.section_num},
                {page: "mod-[a-z]+-mod", mdl_course: {id: this.course_id}},
            );
            this.page_details = await this.page_call({page: "mod-[a-z]+-mod", mdl_course_modules: {course: this.course_id, // section: section_num, modname: "label",
            mdl_course_module_instance: { intro: this.topic_first ?
                `<p></p>

                <div class="header2"> <i class="fa fa-align-justify" aria-hidden="true"></i> ${name}</div>

                <div class="textblock">

                <p>In class, your facilitator will introduce you to
                [introduce the topic here, including learning objectives.
                Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.]
                The online activities listed below support the course work.</p>

                <p><strong>INSTRUCTIONS</strong></p>

                <p>Your facilitator will provide you with information about completing the following activities.
                We suggest that you work through each activity in the sequence set out below, from top to bottom—but
                feel free to complete the activities in the sequence that makes the most sense to you.</p>

                </div>`.replace(/^        /gm, "") :
                `<p></p>

                <div class="header2"> <i class="fa fa-align-justify" aria-hidden="true"></i> ${name}</div>

                <div class="textblock">

                <p>In class, your facilitator will introduce you to
                [introduce the topic here, including learning objectives.
                Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.]
                The online activity listed below supports the course work.</p>

                <p><strong>INSTRUCTIONS</strong></p>

                <p>Your facilitator will provide you with information about completing the following activity.</p>

                </div>`.replace(/^        /gm, ""),
            },
            }, dom_submit: true});
            this.page_details = await this.page_loaded<page_course_view_data>({page: "course-view-[a-z]+", mdl_course: {id: this.course_id}});

            // section = (await ws_call({wsfunction: "core_course_get_contents", courseid: this.course_id, options: [{name: "sectionnumber", value: section_num}]}))[0];
            let section = this.page_details.mdl_course_sections;

            // Move new module.
            this.page_details = await this.page_load("/course/mod.php", {sesskey: this.tabdata.page_sesskey, sr: section.section, copy: section.mdl_course_modules[section.mdl_course_modules.length - 1].id},
            {page: "course-view-[a-z]+"},
            );
            this.page_details = await this.page_load("/course/mod.php", {moveto: this.mod_move_to /*???*/, sesskey: this.tabdata.page_sesskey},
                {page: "course-view-[a-z]+"},
            );

            // Create topic end message (4 loads?)
            if (this.topic_first) {
                this.page_details = await this.page_load("/course/mod.php", {id: this.course_id, sesskey: this.tabdata.page_sesskey, sr: 0 /* TODO: remove? */, add: "label", section: this.section_num},
                    {page: "mod-[a-z]+-mod", mdl_course: {id: this.course_id}},
                );
                this.page_details = await this.page_call({page: "mod-[a-z]+-mod", mdl_course_modules: {course: this.course_id, // section: section_num, modname: "label",
                    mdl_course_module_instance: { intro:
                    `<p></p>

                    <p>When you have completed the above activities, and your facilitator provides you with further information,
                    please continue to the next topic below—<strong>[xxxxxxx]</strong>.</p>`
                    },
                    }, dom_submit: true});
                this.page_details = await this.page_loaded<page_course_view_data>({page: "course-view-[a-z]+", mdl_course: {id: this.course_id}});

                // section = (await ws_call({wsfunction: "core_course_get_contents", courseid: this.course_id, options: [{name: "sectionnumber", value: section_num}]}))[0];
                section = this.page_details.mdl_course_sections;

                // Move new module.
                this.page_details = await this.page_load("/course/mod.php", {sesskey: this.tabdata.page_sesskey, sr: section.section, copy: section.mdl_course_modules[section.mdl_course_modules.length - 1].id},
                {page: "course-view-[a-z]+"},
                );
                this.page_details = await this.page_load("/course/mod.php", {moveto: this.mod_move_to /*???*/, sesskey: this.tabdata.page_sesskey},
                    {page: "course-view-[a-z]+"},
                );
            } else {
                this.tabdata.page_load_count(4);
            }

            // Create space (4 loads?)
            this.page_details = await this.page_load("/course/mod.php", {id: this.course_id, sesskey: this.tabdata.page_sesskey, sr: 0 /* TODO: remove? */, add: "label", section: this.section_num},
                {page: "mod-[a-z]+-mod", mdl_course: {id: this.course_id}},
            );
            this.page_details = await this.page_call({page: "mod-[a-z]+-mod", mdl_course_modules: {course: this.course_id, // section: section_num, modname: "label",
                mdl_course_module_instance: { intro:
                ""
                },
                }, dom_submit: true});
            this.page_details = await this.page_loaded<page_course_view_data>({page: "course-view-[a-z]+", mdl_course: {id: this.course_id}});

            // section = (await ws_call({wsfunction: "core_course_get_contents", courseid: this.course_id, options: [{name: "sectionnumber", value: section_num}]}))[0];
            section = this.page_details.mdl_course_sections;

            // Move new module.
            this.page_details = await this.page_load("/course/mod.php", {sesskey: this.tabdata.page_sesskey, sr: section.section, copy: section.mdl_course_modules[section.mdl_course_modules.length - 1].id},
            {page: "course-view-[a-z]+"},
            );
            this.page_details = await this.page_load("/course/mod.php", {moveto: this.mod_move_to /*???*/, sesskey: this.tabdata.page_sesskey},
                {page: "course-view-[a-z]+"},
            );

        }


    }




    class Test_Macro extends Macro {


        public init(page_details: Page_Data_Out) {
            this.progress_max = 10;
            this.page_details = page_details;
            this.prereq = true;
        }


        protected async content() {

            this.page_details = await this.page_load(
                "/course/index.php", {},
                {page: "course-index(-category)?", mdl_course: {id: 1}},
            );

            // const site_map = await this.page_call({page: "course-index(-category)?", dom_expand: true});

            const course_id = 7015;
            const course_context = 911164;

            await this.page_load(
                "/backup/restorefile.php", {contextid: course_context},
                {page: "backup-restorefile", mdl_course: {id: course_id}},
            );

            // const message = await this.page_call({page: "backup-restorefile"});

            // await browser.downloads.download({url: message.mdl_course.x_backup_url, saveAs: false});

        }


    }




}
