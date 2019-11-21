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

        private static check_with_skel(data: any, skel: any, exclude?: any): boolean {
            if (skel == undefined) { return true; }
            else if (typeof(skel) == "function")  {
                throw new Error("Can't compare functions");
            } else if (typeof(skel) == "object") {
                for (const prop in skel) if (skel.hasOwnProperty(prop) && skel[prop] != undefined) {
                    if (exclude && exclude.hasOwnProperty(prop)) continue;
                    if (!data.hasOwnProperty(prop) || !this.check_with_skel(data[prop], skel[prop])) { throw new Error("Didin't find property: " + prop); }
                }
                return true;
            } else {
                if (data != skel) { throw new Error("Property failed compare."); }
                return true;
            }
        }




        public macro_state:         number = 0;     // -1 error / 0 idle / 1 running / 2 running & awaiting load
        // macro_callstack:    string[3]       // 0: macro  1: macro step  2: tabdata function
        public macro_error:         Errorlike|null = null;
        public macro_progress:      number = 1;
        public macro_progress_max:  number = 1;
        public macro_cancel:        boolean = false;

        public macros: {[index: string] : Macro} = {
            new_course:     new New_Course_Macro(this),
            index_rebuild:  new Index_Rebuild_Macro(this),
            new_section:    new New_Section_Macro(this),
            new_topic:      new New_Topic_Macro(this),
            backup:         new Backup_Macro(this)
        };

        private page_details:       Page_Data_Out|null = null;
        private page_tab_id:        number;

        private page_load_wait:     number = 0;
        private page_message:       Page_Data_Out|Errorlike|null = null;
        private page_is_loaded:     boolean = false;

        private popup:              Popup|null = null;


        constructor(tab_id: number) {
            this.page_tab_id = tab_id;
            void this.init();
        }


        public async init() {
            try {
                if (this.popup) {
                    this.popup.close();
                }
            } catch (e) { }
            this.page_load_wait = 0;
            this.page_message = null;
            this.page_is_loaded = true;
            this.macro_error = null;
            this.macro_progress = 1;
            this.macro_progress_max = 1;
            this.macro_cancel = false;
            try {
                this.macro_state = 1;
                this.page_details = await this.page_call({});
                this.macro_state = 0;
            } catch (e) {
                this.macro_state = 0;
                this.page_details = null;
            }
            this.macros_init(this.page_details);
            this.macro_state = 0;
            this.update_ui();
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
                if (this.popup) {
                    this.popup.update();
                }
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


        public async page_load<T extends Page_Data>(
                                page_data: DeepPartial<T & Page_Data_Out_Base> & {location: {pathname: string, search: {[index: string]: number|string}}},
                                count: number = 1): Promise<T & Page_Data_Out_Base> {
            return await this.page_load2<T>(page_data, page_data, count);
        }


        public async page_load2<T extends Page_Data>(
            page_data1: DeepPartial<Page_Data> & {location: {pathname: string, search: {[index: string]: number|string}}},
            page_data2: DeepPartial<T & Page_Data_Out_Base>,
            count: number = 1
        ): Promise<T & Page_Data_Out_Base> {
            const pathname = page_data1.location.pathname;
            const search = page_data1.location.search;

            (this.macro_state == 1)                                             || throwf(new Error("Page load:\nUnexpected state."));
            (pathname.match(/(?:\/[a-z]+)+\.php/))                              || throwf(new Error("Page load:\nPathname unexpected."));

            this.macro_state = 2;
            this.page_is_loaded = false;
            this.page_message = null;

            await browser.tabs.update(this.page_tab_id, {url: this.page_details.moodle_page.wwwroot + pathname + "?" + TabData.json_to_search_params(search)});
            return await this.page_loaded(page_data2, count);
        }


        public async page_loaded<T extends Page_Data>(page_data: DeepPartial<T & Page_Data_Out_Base>,
            count: number = 1): Promise<T & Page_Data_Out_Base> {

            (this.macro_state == 2)                                             || throwf(new Error("Page loaded:\nUnexpected state."));

            this.page_load_wait  = 0;
            do {
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
                if (this.popup) {
                    this.popup.update_progress();
                }
            } catch (e) { }
        }


        public onMessage(message: Page_Data_Out|Errorlike, _sender: browser.runtime.MessageSender) {

            if (!this.page_message || this.macro_state == 0) {
                this.page_message = message;
                if (is_Errorlike(message)) {
                    this.page_details = null;
                } else {
                    this.page_details = message;
                }

                if (this.page_is_loaded ) {

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
                            }
                        } else {
                            if (this.macro_state == 0) { this.macros_init(this.page_details); }
                            this.update_ui();
                        }
                    }
                } else if (!this.page_message || !is_Errorlike(this.page_message)) {
                    this.page_message = new Error("Unexpected tab update");

                }

                this.update_ui();

            }

        }


        private macros_init(page_details: Page_Data_Out|null) {
            this.macros.new_course.init(page_details);
            this.macros.index_rebuild.init(page_details);
            this.macros.new_section.init(page_details);
            this.macros.new_topic.init(page_details);
            this.macros.backup.init(page_details);
        }


        private page_load_match<T extends Page_Data_Base>(page_details: Page_Data_Base & Page_Data_Out_Base, page_data: DeepPartial<T & Page_Data_Out_Base>):
            page_details is T & Page_Data_Out_Base {
            let result = true;
            if (!page_data.page || page_details.moodle_page.body_id.match(RegExp("^page-" + page_data.page + "$"))) { /* OK */ } else    { result = false; }
            /*for (const prop in body_class) if (body_class.hasOwnProperty(prop)) {
                if ((" " + this.page_details.moodle_page.body_class + " ").match(" " + prop + (body_class[prop] ? ("-" + body_class[prop]) : "") + " "))
                    {  OK  }
                else
                    { result = false; }
            }*/
            // TabData.check_with_skel(page_details/*.moodle_page*/, page_data/*.moodle_page*/, {location: true});
            return result;
        }


    }




    abstract class Macro {


        public prereq:      boolean     = false;

        public params:      {}|null     = null;

        protected tabdata:  TabData;

        protected data:     {}|null     = null;

        protected progress_max: number  = 1;

        protected page_details: Page_Data_Out;

        constructor(new_tabdata: TabData) {
            this.tabdata = new_tabdata;
            this.prereq = false;
        }


        public abstract init(page_details: Page_Data_Out|null): void;


        public async run(): Promise<void> {

            this.params = JSON.parse(JSON.stringify(this.params));

            if (this.tabdata.macro_state != 0) {
                return;
            }

            this.tabdata.macro_cancel   = false;
            this.tabdata.macro_state    = 1;
            this.tabdata.macro_progress = 0;
            this.tabdata.macro_progress_max = this.progress_max;

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

            /*
            this.tabdata.macro_state = 0;
            this.tabdata.macro_progress = this.tabdata.macro_progress_max;
            this.tabdata.macros_init(this.tabdata.page_details);
            try {
                this.tabdata.popup.close();
            } catch (e) {
                // Do nothing
            }
            */
            await this.tabdata.init(); // TODO: Menus sometimes don't display properly after cancel?

        }


        protected abstract async content(): Promise<void>;


    }


    export type New_Course_Params = {
        mdl_course: { fullname: string, shortname: string, startdate: number};
    };

    type New_Course_Data = {
        mdl_course: { template_id: number };
        mdl_course_categories: { id: number; name: string };
    };

    export class New_Course_Macro extends Macro {


        public params: New_Course_Params|null = null;
        protected data: New_Course_Data|null = null;


        public init(page_details: Page_Data_Out) {

            this.prereq = false;

            this.page_details = page_details;

            if (!this.page_details || !this.page_details.hasOwnProperty("page") || this.page_details.page != "course-index(-category)?")     { return; }

            if (this.page_details.moodle_page.body_id != "page-course-index-category") { return; }

            let template_id: number;
            if (page_details.moodle_page.wwwroot == "https://otagopoly-moodle.testing.catlearn.nz" ) {
                template_id = 6548;
            } else if (page_details.moodle_page.wwwroot == "https://moodle.op.ac.nz") {
                template_id = 6548;
            } else                                                                      { return; }

            this.data = {
                mdl_course_categories: this.page_details.mdl_course_categories,
                mdl_course: { template_id: template_id}
            };



            this.progress_max = 17 + 1;
            this.prereq = true;

        }


        protected async content() {  // TODO: Set properties.

            if (!this.data || !this.params)                                     throw new Error("New course macro, prereq:\ndata not set.");

            // Get template course context (1 load)
            this.page_details = await this.tabdata.page_load(
                {location: {pathname: "/course/view.php", search: {id: this.data.mdl_course.template_id, section: 0}},
                page: "course-view-[a-z]+", mdl_course: {id: this.data.mdl_course.template_id}}
            );
            const source_context_match = this.page_details.moodle_page.body_class.match(/(?:^|\s)context-(\d+)(?:\s|$)/)
                                                                                        || throwf(new Error("New course macro, get template:\nContext not found."));
            const source_context = parseInt(source_context_match[1]);

            // Load course restore page (1 load)
            this.page_details = await this.tabdata.page_load(
                {location: {pathname: "/backup/restorefile.php", search: {contextid: source_context}},
                page: "backup-restorefile", mdl_course: {id: this.data.mdl_course.template_id}},
            );

            // Click restore backup file (1 load)
            this.page_details = await this.tabdata.page_call<page_backup_restorefile_data>({page: "backup-restorefile", dom_submit: "restore"});
            this.page_details = await this.tabdata.page_loaded<page_backup_restore_data>({page: "backup-restore", stage: 2, mdl_course: {template_id: this.data.mdl_course.template_id}});

            // Confirm (1 load)
            (this.page_details.stage == 2)                                      || throwf(new Error("New course macro, confirm:\nStage unexpected."));
            this.page_details = await this.tabdata.page_call({page: "backup-restore", stage: 2, dom_submit: "stage 2 submit"});
            this.page_details = await this.tabdata.page_loaded<page_backup_restore_data>({page: "backup-restore", mdl_course: {template_id: this.data.mdl_course.template_id}});

            // Destination: Search for category (1 load)
            (this.page_details.stage == 4)                                      || throwf(new Error("New course macro, destination:\nStage unexpected."));
            this.page_details = await this.tabdata.page_call({page: "backup-restore", stage: 4, mdl_course_categories: {name: this.data.mdl_course_categories.name}, dom_submit: "stage 4 new cat search"});
            this.page_details = await this.tabdata.page_loaded({page: "backup-restore"});  // TODO: Add details

            // Destination: Select category (1 load)
            this.page_details = await this.tabdata.page_call({page: "backup-restore", stage: 4, mdl_course_categories: {id: this.data.mdl_course_categories.id}, dom_submit: "stage 4 new continue"});
            this.page_details = await this.tabdata.page_loaded<page_backup_restore_data>({page: "backup-restore", stage: 4, mdl_course: {template_id: this.data.mdl_course.template_id}});

            // Restore settings (1 load)
            if (this.page_details.stage != 4)                                    { throw new Error("New course macro, restore settings:\nStage unexpected."); }
            this.page_details = await this.tabdata.page_call({page: "backup-restore", stage: 4, restore_settings: {users: false}});
            this.page_details = await this.tabdata.page_call({page: "backup-restore", dom_submit: "stage 4 settings submit"});
            this.page_details = await this.tabdata.page_loaded<page_backup_restore_data>({page: "backup-restore", mdl_course: {template_id: this.data.mdl_course.template_id}});

            // Course settings (1 load)
            (this.page_details.stage == 8)                                      || throwf(new Error("New course macro, course settings:\nStage unexpected."));
            const course = {fullname: this.params.mdl_course.fullname, shortname: this.params.mdl_course.shortname, startdate: this.params.mdl_course.startdate};
            this.page_details = await this.tabdata.page_call({page: "backup-restore", stage: 8, mdl_course: course, dom_submit: "stage 8 submit"});
            this.page_details = await this.tabdata.page_loaded<page_backup_restore_data>({page: "backup-restore", mdl_course: {template_id: this.data.mdl_course.template_id}});

            // Review & Process (~5 loads)
            (this.page_details.stage == 16)                                     || throwf(new Error("New course macro, review & process:\nStage unexpected"));
            this.page_details = await this.tabdata.page_call({page: "backup-restore", stage: 16, dom_submit: "stage 16 submit"});
            this.page_details = await this.tabdata.page_loaded<page_backup_restore_data>({page: "backup-restore", mdl_course: {template_id: this.data.mdl_course.template_id}}, 5);

            // Complete--Go to new course (1 load)
            (this.page_details.stage == null)                                   || throwf(new Error("New course macro, complete:\nStage unexpected."));
            const course_id = this.page_details.mdl_course.id as number;
            this.page_details = await this.tabdata.page_call({page: "backup-restore", stage: null, dom_submit: "stage complete submit"});
            this.page_details = await this.tabdata.page_loaded<page_course_view_data>({page: "course-view-[a-z]+", mdl_course: {id: course_id}});

            // TODO: Turn editing on (1 load).
            if (!this.page_details.moodle_page.body_class.match(/\bediting\b/)) {
                this.page_details = await this.tabdata.page_load<page_course_view_data>(
                    {location: {pathname: "/course/view.php", search: {id: course_id, section: 0, sesskey: this.page_details.moodle_page.sesskey, edit: "on"}},
                    page: "course-view-[a-z]+", mdl_course: {id: course_id}},
                );
            } else {
                this.tabdata.page_load_count(1);
            }
            (this.page_details.moodle_page.body_class.match(/\bediting\b/))     || throwf(new Error("New course macro, turn editing on:\nEditing not on."));

            // Fill in course name (2 loads)
            const section_0_id = this.page_details.mdl_course_sections.id       || throwf(new Error("New course macro, fill in course name:\nID not found."));
            this.page_details = await this.tabdata.page_load<page_course_editsection_data>(
                {location: {pathname: "/course/editsection.php", search: {id: section_0_id}}, // TODO: Needs editing on.
                page: "course-editsection", mdl_course: {id: course_id}},
            );
            const desc = this.page_details.mdl_course_sections.summary.replace(/\[Course Name\]/g, this.params.mdl_course.fullname);
            this.page_details = await this.tabdata.page_call({page: "course-editsection", mdl_course_sections: {summary: desc}, dom_submit: true});
            this.page_details = await this.tabdata.page_loaded({page: "course-view-[a-z]+", mdl_course: {id: course_id}});

        }


    }




    export class Index_Rebuild_Macro extends Macro {


        protected data: {mdl_course: {id: number}; mdl_course_sections: {section: number}; last_section_num: number} | null = null;


        public init(page_details: Page_Data_Out) {

            this.prereq = false;

            this.page_details = page_details;

            // Check course type
            if (!this.page_details || this.page_details.page != "course-view-[a-z]+")                      { return; }
            const course = this.page_details.mdl_course;
            if (!course || course.format != "onetopic" || !course.id) {  return; }

            // Check editing on
            if (!this.page_details.moodle_page || !this.page_details.moodle_page.body_class || !this.page_details.moodle_page.body_class.match(/\bediting\b/)) {
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
            if (modules_tab_num && last_module_tab_num) {  } else                                        {  return; }
            if (this.page_details.mdl_course_sections.section <= last_module_tab_num)
            { } else { return; }

            this.data = {mdl_course: {id: course.id}, mdl_course_sections: {section: modules_tab_num}, last_section_num: last_module_tab_num};

            this.progress_max = last_module_tab_num - modules_tab_num + 3 + 1;
            this.prereq = true;

        }


        protected async content() {

            if (!this.data /*|| !this.params*/)                                     throw new Error("Index rebuild macro, prereq:\ndata not set.");
            // TODO: Don't include hidden tabs or topic headings?

            const parser = new DOMParser();

            // Get list of sections (1 load)
            this.page_details = await this.tabdata.page_load<page_course_view_data>(
                {location: {pathname: "/course/view.php", search: {id: this.data.mdl_course.id, section: this.data.mdl_course_sections.section}},
                page: "course-view-[a-z]+", mdl_course: {id: this.data.mdl_course.id}},
            );
            const modules_list = this.page_details.mdl_course_sections.mdl_course_modules;
            (modules_list.length == 1)                                          || throwf(new Error("Index rebuild macro, get list:\nExpected exactly one resource in Modules tab."));
            const modules_index = modules_list[0];


            // Get section contents (1 load per section)
            let index_html = '<div class="textblock">\n';
            // modules_list.shift();
            for (let section_num = this.data.mdl_course_sections.section + 1; section_num <= this.data.last_section_num; section_num++) {
                // const section_num = section.section                                     || throwf(new Error("Module number not found."));
                this.page_details = await this.tabdata.page_load<page_course_view_data>(
                    {location: {pathname: "/course/view.php", search: {id: this.data.mdl_course.id, section: section_num}},
                    page: "course-view-[a-z]+", mdl_course: {id: this.data.mdl_course.id}});
                const section_full = this.page_details.mdl_course_sections;
                const section_name = (parser.parseFromString(section_full.summary as string, "text/html").querySelector(".header1")
                                                                                        || throwf(new Error("Index rebuild macro, get section:\nModule name not found."))
                                    ).textContent                                      || throwf(new Error("Index rebuild macro, get section:\nModule name content not found."));
                index_html = index_html
                            + '<a href="' + this.page_details.moodle_page.wwwroot + "/course/view.php?id=" + this.data.mdl_course.id + "&section=" + section_num + '"><b>' + TabData.escapeHTML(section_name.trim()) + "</b></a>\n"
                            + "<ul>\n";
                for (const mod of section_full.mdl_course_modules) {
                    // parse description
                    const mod_desc = parser.parseFromString((mod.mdl_course_module_instance).intro || "", "text/html");
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
            this.page_details = await this.tabdata.page_load2(
                {location: {pathname: "/course/mod.php", search: {sesskey: this.page_details.moodle_page.sesskey, update: modules_index.id}}},
                {page: "mod-[a-z]+-mod", mdl_course: {id: this.data.mdl_course.id}});
            this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_modules: {course: this.data.mdl_course.id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
                mdl_course_module_instance: {intro: index_html}}, dom_submit: true});
            this.page_details = await this.tabdata.page_loaded({page: "course-view-[a-z]+", mdl_course: {id: this.data.mdl_course.id}});

        }


    }




    export class New_Section_Macro extends Macro {


        public params: {mdl_course_sections: {name: string, fullname: string}}|null = null;
        protected data: {mdl_course: {id: number}, mdl_course_sections: {section: number}, mdl_course_modules: {feedback_template_id: number}}|null = null;


        public init(page_details: Page_Data_Out) {

            this.prereq = false;
            // Get site details
            // TODO: Also customise image link per site

            this.page_details = page_details;

            // Check page type
            if (!this.page_details || this.page_details.page != "course-view-[a-z]+") {
                return;
            }

            let feedback_template_id: number;
            if (this.page_details.moodle_page.wwwroot == "https://otagopoly-moodle.testing.catlearn.nz" ) {
                feedback_template_id = 59;
            } else if (this.page_details.moodle_page.wwwroot == "https://moodle.op.ac.nz") {
                feedback_template_id = 59;
            } else                                                                      { return; }

            // Check editing on
            if (!this.page_details.moodle_page || !this.page_details.moodle_page.body_class || !this.page_details.moodle_page.body_class.match(/\bediting\b/)) {
                return;
            }

            // Get course details
            const course = this.page_details.mdl_course;
            if (!course) { return; }

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
            if (modules_tab_num && last_module_tab_num) {  } else                                        { return; }
            if (this.page_details.mdl_course_sections.section <= last_module_tab_num)
            { } else { return; }
            // this.new_section_pos = last_module_tab_num + 1;

            this.data = {mdl_course: {id: course.id}, mdl_course_sections: {section: last_module_tab_num + 1}, mdl_course_modules: {feedback_template_id: feedback_template_id}};

            this.progress_max = 17 + 1;
            this.prereq = true;
        }


        protected async content() {

            if (!this.data || !this.params)                                     throw new Error("New section macro, prereq:\ndata not set.");

            // Add new tab (1 load)
            this.page_details = await this.tabdata.page_load2<page_course_view_data>(  // TODO: Fix for flexsections?
                {location: {pathname: "/course/changenumsections.php", search: {courseid: this.data.mdl_course.id, increase: 1, sesskey: this.page_details.moodle_page.sesskey, insertsection: 0}}},
                {page: "course-view-[a-z]+", mdl_course: {id: this.data.mdl_course.id}},
            );
            let new_section = this.page_details.mdl_course_sections;

            // Move new tab (1 load)
            this.page_details = await this.tabdata.page_load(
                {location: {pathname: "/course/view.php", search: {id: this.data.mdl_course.id, section: new_section.section, sesskey: this.page_details.moodle_page.sesskey, move: this.data.mdl_course_sections.section - new_section.section},
                                                                                    /*|| throwf(new Error("WS course section edit, no amount specified."))*/},
                page: "course-view-[a-z]+", mdl_course: {id: this.data.mdl_course.id}},
            );
            new_section.section = this.data.mdl_course_sections.section;

            // Set new tab details (2 loads)
            this.page_details = await this.tabdata.page_load(
                {location: {pathname: "/course/editsection.php", search: {id: new_section.id, sr: new_section.section,
                                                                                    /*|| throwf(new Error("WS course section edit, no amount specified."))*/}},
                                                                                    page: "course-editsection", mdl_course: {id: this.data.mdl_course.id}},
                                                                                    );
            this.page_details = await this.tabdata.page_call({page: "course-editsection", mdl_course_sections: {id: new_section.id, name: this.params.mdl_course_sections.name, x_options: {level: 1},
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
            this.page_details = await this.tabdata.page_loaded({page: "course-view-[a-z]+", mdl_course: {id: this.data.mdl_course.id}});

            // Add pre-topic message (2 loads)
            // Body ID or class unexpected.
            this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {id: this.data.mdl_course.id, sesskey: this.page_details.moodle_page.sesskey, sr: new_section.section, add: "label", section: new_section.section}}},
                                    {page: "mod-[a-z]+-mod", mdl_course: {id: this.data.mdl_course.id}},
                                    );
            this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_modules: {course: this.data.mdl_course.id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
                mdl_course_module_instance: {intro:
                `<p></p>

                <p>After you have worked through all of the above topics, and your facilitator provides you with further information in class,
                you're now ready to demonstrate evidence of what you have learnt in this module.
                Please click on the <strong>Assessments</strong> tab above for further information.</p>`.replace(/^        /gm, "")},
                }, dom_submit: true});
            this.page_details = await this.tabdata.page_loaded({page: "course-view-[a-z]+", mdl_course: {id: this.data.mdl_course.id}});

            // Add blank line (2 loads)
            this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {id: this.data.mdl_course.id, sesskey: this.page_details.moodle_page.sesskey, sr: new_section.section, add: "label", section: new_section.section}}},
            {page: "mod-[a-z]+-mod", mdl_course: {id: this.data.mdl_course.id}},
            );
            this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_modules: {course: this.data.mdl_course.id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
            mdl_course_module_instance: {intro: ""}, }, dom_submit: true});
            this.page_details = await this.tabdata.page_loaded({page: "course-view-[a-z]+", mdl_course: {id: this.data.mdl_course.id}});

            // Add feedback topic (2 loads)
            this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {id: this.data.mdl_course.id, sesskey: this.page_details.moodle_page.sesskey, sr: new_section.section, add: "label", section: new_section.section}}},
                {page: "mod-[a-z]+-mod", mdl_course: {id: this.data.mdl_course.id}},
            );
            this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_modules: {course: this.data.mdl_course.id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
            mdl_course_module_instance: {intro:
                `<p></p>

                <p><strong>YOUR FEEDBACK</strong></p>

                <p>We appreciate your feedback about your experience with working through this module.
                Please click the 'Your feedback' link below if you wish to respond to a five-question survey.
                Thanks!</p>`.replace(/^        /gm, "")},
            }, dom_submit: true});
            this.page_details = await this.tabdata.page_loaded({page: "course-view-[a-z]+", mdl_course: {id: this.data.mdl_course.id}});

            // Add feedback activity (2 load)
            this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {id: this.data.mdl_course.id, sesskey: this.page_details.moodle_page.sesskey, sr: new_section.section, add: "feedback", section: new_section.section}}},
                {page: "mod-[a-z]+-mod", mdl_course: {id: this.data.mdl_course.id}},
            );
            this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_modules: {course: this.data.mdl_course.id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
            mdl_course_module_instance: {name: "Your feedback", intro:
                `<div class="header2"> <i class="fa fa-bullhorn" aria-hidden="true"></i> FEEDBACK</div>

                <div class="textblock">

                <p><strong>DESCRIPTION</strong></p>

                <p>Please help us improve this learning module by answering five questions about your experience.
                This survey is anonymous.</p>
                </div>`.replace(/^        /gm, "")}, }, dom_submit: true});
            this.page_details = await this.tabdata.page_loaded<page_course_view_data>({page: "course-view-[a-z]+", mdl_course: {id: this.data.mdl_course.id}});
            new_section = this.page_details.mdl_course_sections;
            let feedback_act: page_course_view_course_modules|null = null;
                for (const module of new_section.mdl_course_modules) {
                    if (!feedback_act || module.id > feedback_act.id) {
                        feedback_act = module;
                    }
                }

            // Configure Feedback activity (3 loads?)
            this.page_details = await this.tabdata.page_load({location: {pathname: "/mod/feedback/edit.php", search: {id: feedback_act.id, do_show: "templates"}},
                                page: "mod-feedback-edit", mdl_course_modules: {id: feedback_act.id}});
            this.page_details = await this.tabdata.page_call({page: "mod-feedback-edit", mdl_course_modules: { mdl_course_module_instance: {mdl_feedback_template_id: this.data.mdl_course_modules.feedback_template_id}}, dom_submit: true});  // TODO: fix;
            this.page_details = await this.tabdata.page_loaded({page: "mod-feedback-use_templ"});
            this.page_details = await this.tabdata.page_call({page: "mod-feedback-use_templ", dom_submit: true});
            this.page_details = await this.tabdata.page_loaded({page: "mod-feedback-edit"});

            // Add footer (2 loads).
            this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {id: this.data.mdl_course.id, sesskey: this.page_details.moodle_page.sesskey, sr: new_section.section, add: "label", section: new_section.section}}},
            {page: "mod-[a-z]+-mod", mdl_course: {id: this.data.mdl_course.id}},
            );
            this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_modules: {course: this.data.mdl_course.id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
            mdl_course_module_instance: {intro:
                `<p></p>

                <p><span style="font-size: xx-small;">
                Image: <a href="https://stock.tookapic.com/photos/12801" target="_blank">Blooming</a>
                by <a href="https://stock.tookapic.com/pawelkadysz" target="_blank">Paweł Kadysz</a>,
                licensed under <a href="https://creativecommons.org/publicdomain/zero/1.0/deed.en" target="_blank">CC0</a>
                </span></p>`.replace(/^        /gm, "")},
            }, dom_submit: true});
            this.page_details = await this.tabdata.page_loaded({page: "course-view-[a-z]+", mdl_course: {id: this.data.mdl_course.id}});

        }


    }




    export class New_Topic_Macro extends Macro {


        public params: {mdl_course_modules: {fullname: string}}|null = null;

        protected data: {mdl_course: {id: number}, mdl_course_sections: {section: number}, mdl_course_modules: {next_id: number, first_topic: boolean}}|null = null;

        public init(page_details: Page_Data_Out) {

            this.prereq = false;

            this.page_details = page_details;

            if (!this.page_details || this.page_details.page != "course-view-[a-z]+")                      { return; }
            const course = this.page_details.mdl_course;
            if (course && course.hasOwnProperty("format") && course.format == "onetopic" && course.id) {  } else { return; }


            // Check editing on
            if (!this.page_details.moodle_page || !this.page_details.moodle_page.body_class || !this.page_details.moodle_page.body_class.match(/\bediting\b/)) {
                return;
            }

            const section = this.page_details.mdl_course_sections;

            let mod_pos = section.mdl_course_modules.length - 1;
            let mod_match_pos = 3;

            while (mod_pos > -1 && mod_match_pos > -1) {
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
                }
            }

            if (mod_match_pos < 0) {  } else                                      { return; }

            const mod_move_to = section.mdl_course_modules[mod_pos + 1].id;

            const topic_first = (mod_pos < 0) ? true : false;

            this.data = {mdl_course: {id: course.id}, mdl_course_sections: {section: section.section}, mdl_course_modules: {next_id: mod_move_to, first_topic: topic_first}};

            this.progress_max = 12 + 1;
            this.prereq = true;
        }


        protected async content() {

            if (!this.data || !this.params)                                     throw new Error("New section macro, prereq:\ndata not set.");

            const name = this.params.mdl_course_modules.fullname;

            // Create topic heading (4 loads?)
            this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {id: this.data.mdl_course.id, sesskey: this.page_details.moodle_page.sesskey, sr: 0 /* TODO: remove? */, add: "label", section: this.data.mdl_course_sections.section}}},
                {page: "mod-[a-z]+-mod", mdl_course: {id: this.data.mdl_course.id}},
            );
            this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_modules: {course: this.data.mdl_course.id, // section: section_num, modname: "label",
            mdl_course_module_instance: { intro: this.data.mdl_course_modules.first_topic ?
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
            this.page_details = await this.tabdata.page_loaded<page_course_view_data>({page: "course-view-[a-z]+", mdl_course: {id: this.data.mdl_course.id}});

            let section = this.page_details.mdl_course_sections;

            // Move new module.
            this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {sesskey: this.page_details.moodle_page.sesskey, sr: section.section, copy: section.mdl_course_modules[section.mdl_course_modules.length - 1].id}}},
                {page: "course-view-[a-z]+"},
            );
            this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {moveto: this.data.mdl_course_modules.next_id /*???*/, sesskey: this.page_details.moodle_page.sesskey}}},
                {page: "course-view-[a-z]+"},
            );

            // Create topic end message (4 loads?)
            if (this.data.mdl_course_modules.first_topic) {
                this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {id: this.data.mdl_course.id, sesskey: this.page_details.moodle_page.sesskey, sr: 0 /* TODO: remove? */, add: "label", section: this.data.mdl_course_sections.section}}},
                    {page: "mod-[a-z]+-mod", mdl_course: {id: this.data.mdl_course.id}},
                );
                this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_modules: {course: this.data.mdl_course.id, // section: section_num, modname: "label",
                    mdl_course_module_instance: { intro:
                    `<p></p>

                    <p>When you have completed the above activities, and your facilitator provides you with further information,
                    please continue to the next topic below—<strong>[xxxxxxx]</strong>.</p>`
                    },
                    }, dom_submit: true});
                this.page_details = await this.tabdata.page_loaded<page_course_view_data>({page: "course-view-[a-z]+", mdl_course: {id: this.data.mdl_course.id}});

                section = this.page_details.mdl_course_sections;

                // Move new module.
                this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {sesskey: this.page_details.moodle_page.sesskey, sr: section.section, copy: section.mdl_course_modules[section.mdl_course_modules.length - 1].id}}},
                    {page: "course-view-[a-z]+"},
                );
                this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {moveto: this.data.mdl_course_modules.next_id /*???*/, sesskey: this.page_details.moodle_page.sesskey}}},
                    {page: "course-view-[a-z]+"},
                );
            } else {
                this.tabdata.page_load_count(4);
            }

            // Create space (4 loads?)
            this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {id: this.data.mdl_course.id, sesskey: this.page_details.moodle_page.sesskey, sr: 0 /* TODO: remove? */, add: "label", section: this.data.mdl_course_sections.section}}},
                {page: "mod-[a-z]+-mod", mdl_course: {id: this.data.mdl_course.id}},
            );
            this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_modules: {course: this.data.mdl_course.id, // section: section_num, modname: "label",
                mdl_course_module_instance: { intro:
                ""
                },
                }, dom_submit: true});
            this.page_details = await this.tabdata.page_loaded<page_course_view_data>({page: "course-view-[a-z]+", mdl_course: {id: this.data.mdl_course.id}});

            section = this.page_details.mdl_course_sections;

            // Move new module.
            this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {sesskey: this.page_details.moodle_page.sesskey, sr: section.section, copy: section.mdl_course_modules[section.mdl_course_modules.length - 1].id}}},
            {page: "course-view-[a-z]+"},
            );
            this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {moveto: this.data.mdl_course_modules.next_id /*???*/, sesskey: this.page_details.moodle_page.sesskey}}},
                {page: "course-view-[a-z]+"},
            );

        }


    }




    export class Backup_Macro extends Macro {

        public params: { }; // mdl_course_categories: { mdl_course: {id: number}[]} };

        public init(page_details: Page_Data_Out) {
            this.prereq     = false;
            if (!page_details || page_details.page != "course-management") { return; }
            this.progress_max = 1000;
            this.page_details = page_details;
            this.data       = {};
            this.prereq     = true;
        }


        private static expand_ticked(category: page_course_management_category, parent_ticked?: boolean): boolean {
            let change: boolean = false;
            if (parent_ticked && !category.checked) {
                category.checked = true;
                change = true;
            }
            if ((category.checked || parent_ticked) && category.expandable && !category.expanded) {
                category.expanded = true;
                change = true;
            }
            for (const subcategory of category.mdl_course_categories) {
                change = change || this.expand_ticked(subcategory, parent_ticked || category.checked);
            }
            return change;
        }

        private static ticked_categories(category: page_course_management_category): page_course_management_category[] {
            let result: page_course_management_category[] = [];
            for (const subcategory of category.mdl_course_categories) {
                if (subcategory.checked) {
                    result.push(subcategory);
                }
                if (subcategory.mdl_course_categories.length > 0) {
                    result = result.concat(Backup_Macro.ticked_categories(subcategory));
                }
            }
            return result;
        }
    

        protected async content() {

            //this.progress_max = this.params.mdl_course_categories.mdl_course.length * 12 + 1;
            //this.tabdata.macro_progress_max = this.progress_max;

            // const site_map = await this.tabdata.page_call({page: "course-index(-category)?", dom_expand: true});
            let change: boolean;
            do {
                this.page_details = await this.tabdata.page_call<page_course_management_data>({page: "course-management"});
                let site_map: page_course_management_category = this.page_details.mdl_course_categories;
                change = Backup_Macro.expand_ticked(site_map);
                if (change) {
                    this.page_details = await this.tabdata.page_call<page_course_management_data>({page: "course-management", mdl_course_categories: site_map});
                    //await sleep(1000);
                }
            } while (change)

            const category_list = Backup_Macro.ticked_categories(this.page_details.mdl_course_categories);

            // Calculate max progress
            let course_count = 0;
            for (const category of category_list) {
                course_count += category.coursecount;
            }
            this.progress_max = category_list.length + course_count * 12 + 1;
            this.tabdata.macro_progress_max = this.progress_max;

            // Get course list from category list
            let course_list: page_course_management_course[] = [];
            for (const category of category_list) {
                this.page_details = await this.tabdata.page_load<page_course_management_data>({location: {pathname: "/course/management.php", search: {categoryid: category.id, perpage: 999}}});
                course_list = course_list.concat(this.page_details.mdl_course);
            }

            for (const course of course_list) { // this.params.mdl_course_categories.mdl_course) {

                // Create backup file (9 loads)
                this.page_details = await this.tabdata.page_load({location: {pathname: "/backup/backup.php", search: {id: course.id}}, page: "backup-backup"});
                // const course_context_match = this.page_details.moodle_page.body_class.match(/(?:^|\s)context-(\d+)(?:\s|$)/)
                //                                                                || throwf(new Error("Backup macro, create backup:\nContext not found."));
                // const course_context = parseInt(course_context_match[1]);

                this.page_details = await this.tabdata.page_call({page: "backup-backup", dom_submit: "next"});
                this.page_details = await this.tabdata.page_loaded({page: "backup-backup"});

                this.page_details = await this.tabdata.page_call({page: "backup-backup", dom_submit: "next"});
                this.page_details = await this.tabdata.page_loaded<page_backup_backup_data>({page: "backup-backup"});
                const backup_filename = this.page_details.backup.filename;

                this.page_details = await this.tabdata.page_call({page: "backup-backup", dom_submit: "perform backup"});
                this.page_details = await this.tabdata.page_loaded({page: "backup-backup"}, 5);

                this.page_details = await this.tabdata.page_call({page: "backup-backup", dom_submit: "continue"});
                this.page_details = await this.tabdata.page_loaded<page_backup_restorefile_data>({page: "backup-restorefile"});

                /*
                this.page_details = await this.tabdata.page_load(
                    {location: {pathname: "/backup/restorefile.php", search: {contextid: course_context}},
                    page: "backup-restorefile", mdl_course: {id: course_id}},
                );
                */

                // Download backup file (1 load?)
                const backup_index: number = this.page_details.mdl_course.backups.findIndex(function(value) { return value.filename == backup_filename; });
                const backup_url: string = this.page_details.mdl_course.backups[backup_index].download_url;
                const backup_download_id = await browser.downloads.download({url: backup_url, saveAs: false});
                let backup_download_state: string;
                do {
                    await sleep(100);
                    backup_download_state = (await browser.downloads.search({id: backup_download_id}))[0].state;
                } while (backup_download_state == "in_progress");
                if (backup_download_state != "complete") { throw new Error("Download error"); }
                this.tabdata.page_load_count(1);

                // Delete backup file (2 loads?)
                this.page_details = await this.tabdata.page_call({page: "backup-restorefile", dom_submit: "manage"});
                this.page_details = await this.tabdata.page_loaded({page: "backup-backupfilesedit"});

                this.page_details = await this.tabdata.page_call({page: "backup-backupfilesedit", mdl_course: { backups: [{filename: backup_filename, click: true}]}});

                this.page_details = await this.tabdata.page_call({page: "backup-backupfilesedit", backup: {click: "delete"}});
                this.page_details = await this.tabdata.page_call({page: "backup-backupfilesedit", backup: {click: "delete_ok"}});
                this.page_details = await this.tabdata.page_call({page: "backup-backupfilesedit", dom_submit: "save"});
                this.page_details = await this.tabdata.page_loaded({page: "backup-restorefile"});


            }
        }


    }




}
