/**
 * Moodle JS Macro Scripts
 */


namespace MJS {


    export class TabData {


        page_tab_id:        number;

        //page_details_or_error: Page_Data_Out|Errorlike;
        page_details:       Page_Data_Out;
        page_wwwroot:       string;
        page_sesskey:       string;

        page_is_loaded:     boolean = false;
        page_message:       Page_Data_Out|Errorlike = null;
        //page_load_state:    number = 0;     // 0 loading / 1 script loaded / 2 page loaded / 3 script & page loaded
        page_load_wait:     number = 0;

        macro_state:        number = 0;     // -1 error / 0 idle / 1 running / 2 awaiting load / 3 loaded
        macro_error:        Errorlike;
        macro_progress:     number = 100;
        macro_progress_max: number = 100;
        macro_cancel:       boolean = false;

        popup:              Popup;


        macros: {[index:string] : Macro} = {
            new_course: new New_Course_Macro(this),
            index_rebuild: new Index_Rebuild_Macro(this),
            new_section: new New_Section_Macro(this),
            new_topic: new New_Topic_Macro(this),
            test: new Test_Macro(this)
        };


        private static json_to_search_params(search_json: {[index: string]: number|string}): string {
            let search_obj: URLSearchParams = new URLSearchParams();
            for (const key in search_json)
                if (search_json.hasOwnProperty(key) && search_json[key] != undefined) {
                    search_obj.append(key, "" + search_json[key]);
            }
            return search_obj.toString();
        }


        static escapeHTML(text: string) {
            return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");  // TODO: Line breaks?  nbsp?
        }


        constructor(tab_id: number) {
            this.page_tab_id = tab_id;
            this.macro_state = 0;
            //this.page_load_state = 0;
            this.page_is_loaded = false;
            this.page_message = null;
            //console.log("load state: "+ this.page_load_state);
        }





        public async page_call(message: Page_Data_In): Promise<Page_Data_Out> {

            (this.macro_state == 1)                                             || throwf(new Error("MJS page call: Unexpected state."));

            if (message.dom_submit) {
                this.macro_state = 2;
            }
            const result = await browser.tabs.sendMessage(this.page_tab_id, message) as Page_Data_Out|Errorlike;
            (!is_Errorlike(result))                                            || throwf(new Error(result.message));

            return result as Page_Data_Out;
        }


        public async page_load(pathname: string, search: {[index: string]: number|string},
                                body_id_start: string, body_class: {[index: string]: string|number},
                                count: number = 1): Promise<void> {

            (this.macro_state == 1)                                             || throwf(new Error("MJS page load: Unexpected state."));
            (pathname.match(/(?:\/[a-z]+)+\.php/))                              || throwf(new Error("MJS page load: Pathname unexpected."));

            this.macro_state = 2;

            await browser.tabs.update(this.page_tab_id, {url: this.page_wwwroot + pathname + "?" + TabData.json_to_search_params(search)});
            await this.page_loaded(body_id_start, body_class, count);
        }


        public async page_loaded(body_id_start: string, body_class: {[index: string]: string|number}, count: number = 1): Promise<void> {

            (this.macro_state == 2)                                             || throwf(new Error("MJS page loaded: Unexpected state."));

            this.page_load_wait  = 0;
            do {
                //(this.page_load_wait < 600)                                     || throwf(new Error("MJS page loaded: Timed out."));
                await sleep(100);
                if (this.macro_cancel)                                          { throw new Error("Cancelled"); }
                this.page_load_wait += 1;
                if (this.page_load_wait <= count * 30) {  // Assume a step takes 3 seconds
                    this.page_load_count(1 / 30);
                }
            } while (/*this.macro_state != 3*/!this.page_is_loaded || !this.page_message);
            this.macro_state = 3;
            (!is_Errorlike(this.page_message))                        || throwf(new Error(this.page_message.message));
            this.page_details = this.page_message as Page_Data_Out;
            if (this.page_load_wait <= count * 30) {
                this.page_load_count(count - this.page_load_wait / 30);
            }
            (this.page_load_match(body_id_start, body_class))                   || throwf(new Error("MJS page loaded: Body ID or class unexpected."));
            this.macro_state = 1;
        }


        public page_load_count(count: number = 1): void {
            this.macro_progress += count;
        }


        private page_load_match(body_id_start: string, body_class: {[index: string]: string|number}): boolean {
            let result = true;
            if (this.page_details.page_window.body_id.startsWith(body_id_start)) { /* OK */ } else    { result = false; }
            for (const prop in body_class) if (body_class.hasOwnProperty(prop)) {
                if ((" " + this.page_details.page_window.body_class + " ").match(" " + prop + (body_class[prop] ? ("-" + body_class[prop]) : "") + " "))
                    { /* OK */ }
                else
                    { result = false; }
            }
            return result;
        }


        public async onMessage(message: Page_Data_Out|Errorlike, _sender: browser.runtime.MessageSender) {
            //this.page_details_or_error = message;

            if (!this.page_message) {
                this.page_message = message;
                if (is_Errorlike(message)) {
                    this.page_wwwroot = null;
                    this.page_sesskey = null;
                } else {
                    this.page_details = message;
                    console.log("updated page details");
                    this.page_wwwroot = this.page_details.page_window.location_origin;
                    this.page_sesskey = this.page_details.page_window.sesskey;
                }

                if (this.page_is_loaded ) {
                    //this.page_load_state = 3;

                    console.log("*** late page message ***");
                    //console.log("load state: "+ this.page_load_state);
                    //if (this.macro_state==2) {/*this.macro_state = 3;*/}
                    /*else*/ if (this.macro_state==0) { this.macros_init(); }
                    try {
                        await this.popup.update();
                    } catch(e) {
                        // Do nothing
                    }
                }
            } else if (!is_Errorlike(this.page_message)) {
                //console.log("Unexpected state: " + this.page_load_state);
                this.page_message = {name:"Error", message: "Duplicate message"};
            }

        }


        public async onTabUpdated(_tab_id: number, update_info: Partial<browser.tabs.Tab>, _tab: browser.tabs.Tab): Promise<void> {
            //this.m_log += "tab updated "+ _update_info.status +"\n";
            //if (this.m_tab && tab_id == this.m_tab.id) {
                //p_update();
            //}
            if (update_info && update_info.status) {
                if (update_info.status == "loading") {
                    //this.page_load_state = 0;
                    this.page_is_loaded = false;
                    this.page_message = null;
                    //console.log("load state: "+ this.page_load_state);
                    //(this.macro_state == 0 || this.macro_state == 2) || throwf(new Error("unexpected state: " /*+ this.page_load_state*/ + " for status: " + update_info.status));
                } else if (update_info.status == "complete" && /*this.page_load_state == 0*/ !this.page_is_loaded) {
                    this.page_is_loaded = true;
                    if (!this.page_message) {
                        //plugin_started = true;
                        //this.page_load_state = this.page_load_state + 2;

                        //console.log("load state: "+ this.page_load_state);
                        //(this.macro_state == 0 || this.macro_state == 2) || throwf(new Error("unexpected state: " /*+ this.page_load_state*/ + " for status: " + update_info.status));
                        if (this.macro_state == 2) { console.log("*** missing page message ***"); }
                    } else {
                        //this.page_load_state = this.page_load_state + 2;
                        //console.log("load state: "+ this.page_load_state);
                        if (this.macro_state==0) { this.macros_init(); }
                        //else if (this.macro_state==2) {/*this.macro_state = 3;*/}
                        //else { throw new Error("unexpected state: " /*+ this.page_load_state*/ + " for status: " + update_info.status); }
                        try {
                            await this.popup.update();
                        } catch(e) {
                            // Do nothing
                        }
                    }
                } else {
                    //this.m_state = -1;
                    throw new Error("unexpected state: " /*+ this.page_load_state*/ + " for status: " + update_info.status);
                }
            }

            try {
                await this.popup.update();
            } catch(e) {
                // Do nothing
            }
        }


        macros_init() {
            this.macros.new_course.init();
            this.macros.index_rebuild.init();
            this.macros.new_section.init();
            this.macros.new_topic.init();
        }









    }


    abstract class Macro {

        protected tabdata: TabData;


        constructor(new_tabdata: TabData) {
            this.tabdata = new_tabdata;
            this.prereq = false;
        }

        protected get page_details(): Page_Data_Out {
            return this.tabdata.page_details;
        }

        protected set page_details(page_details_in: Page_Data_Out) {
            (page_details_in == this.tabdata.page_details)                      || throwf(new Error("Page details mismatch."));
        }

        protected async page_call(message: Page_Data_In): Promise<Page_Data_Out> {
            return await this.tabdata.page_call(message);
        }

        protected async page_load(pathname: string, search: {[index: string]: number|string},
            body_id_start: string, body_class: {[index: string]: string|number},
            count: number = 1): Promise<void> {
            await this.tabdata.page_load(pathname, search, body_id_start, body_class, count);
        }

        protected async page_loaded(body_id_start: string, body_class: {[index: string]: string|number}, count: number = 1): Promise<void> {
            await this.tabdata.page_loaded(body_id_start, body_class, count);
        }


        prereq: boolean = false;

        abstract init(): void;

        progress_max: number;

        async run(): Promise<void> {

            if (this.tabdata.macro_state != 0) {
                return;
            }

            this.init();

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
                    this.tabdata.popup.update();
                    return;
                }
            }

            this.tabdata.macro_state = 0;
            this.tabdata.macro_progress = this.tabdata.macro_progress_max;
            this.tabdata.macros_init();
            try {
                this.tabdata.popup.close();
            } catch(e) {
                // Do nothing
            }

        };

        abstract async content(): Promise<void>;







    }



    export class New_Course_Macro extends Macro {

        prereq:             boolean;

        private course_template_id: number;
        private category_id:        number;
        private category_name:      string;

        new_course: DeepPartial<MDL_Course> & {
            fullname: string;
            shortname: string;
            startdate: number;
        }


        public init() {
            
            console.log("new course pre starting");
            this.prereq = false;

            if (this.page_details.page != "course-index-category?")     { return; }

            if (this.page_details.page_window.body_id != "page-course-index-category") { return; }

            this.category_id  = this.page_details.mdl_course_categories.id;
            this.category_name = this.page_details.mdl_course_categories.name;

            if (this.tabdata.page_wwwroot == "https://otagopoly-moodle.testing.catlearn.nz" ) {
                this.course_template_id = 6548;
            } else if (this.tabdata.page_wwwroot == "https://moodle.op.ac.nz") {
                this.course_template_id = 6548;
            } else                                                                      { return; }

            this.progress_max = 14;
            this.prereq = true;
        
        }

        public async content() {  // TODO: Set properties.

            // Start
            //this.macro_start();

            const name = this.new_course.fullname;
            const shortname = this.new_course.shortname;
            const startdate = this.new_course.startdate;



            // Get template course context (1 load)
            await this.page_load("/course/view.php", {id: this.course_template_id, section: 0},
                "page-course-view-", {course: this.course_template_id},
            );
            const source_context_match = this.page_details.page_window.body_class.match(/(?:^|\s)context-(\d+)(?:\s|$)/)
                                                                                        || throwf(new Error("WS course restore, source context not found."));
            const source_context = parseInt(source_context_match[1]);

            // Load course restore page (1 load)
            await this.page_load(
                "/backup/restorefile.php", {contextid: source_context},
                "page-backup-restorefile", {course: this.course_template_id},
            );

            // Click restore backup file (1 load)
            await this.page_call({page: "backup-restorefile", dom_submit: "restore"});
            await this.page_loaded("page-backup-restore", {course: this.course_template_id});
            this.page_details = this.tabdata.page_details;
            if (this.page_details.page != "backup-restore")                       { throw (new Error("Click restore backup file, page response unexpected.")) };

            // Confirm (1 load)
            (this.page_details.stage == 2)                                      || throwf(new Error("WS course_restore, step 1 state unexpected."));
            await this.page_call({page: "backup-restore", dom_submit: "stage 2 submit"});
            await this.page_loaded("page-backup-restore", {course: this.course_template_id});

            // Destination: Search for category (1 load)
            (this.page_details.stage == 4)                                      || throwf(new Error("WS course_restore, step 2 state unexpected."));
            await this.page_call({mdl_course_categories: {name: this.category_name}, dom_submit: "stage 4 new cat search"});
            await this.page_loaded("", {});  // TODO: Add details
            
            // Destination: Select category (1 load)
            await this.page_call({mdl_course_categories: {id: this.category_id}, dom_submit: "stage 4 new continue"});       
            await this.page_loaded("page-backup-restore", {course: this.course_template_id});

            // Restore settings (1 load)
            (this.page_details.stage == 4)                                      || throwf(new Error("WS course_restore, step 3 state unexpected."));
            await this.page_call({dom_set_key: "stage 4 settings users", dom_set_value: false});
            await this.page_call({dom_submit: "stage 4 settings submit"});
            await this.page_loaded("page-backup-restore", {course: this.course_template_id});

            // Course settings (1 load)
            (this.page_details.stage == 8)                                      || throwf(new Error("WS course_restore, step 4 state unexpected."));
            const course: DeepPartial<MDL_Course> = {fullname: name, shortname: shortname, startdate: startdate};
            await this.page_call({mdl_course: course, dom_submit: "stage 8 submit"});
            await this.page_loaded("page-backup-restore", {course: this.course_template_id});

            // Review & Process (~5 loads)
            (this.page_details.stage == 16)                                     || throwf(new Error("WS course_restore, step 5 state unexpected."));
            await this.page_call({mdl_course: course, dom_submit: "stage 16 submit"});
            await this.page_loaded("page-backup-restore", {course: this.course_template_id}, 5);

            // Complete--Go to new course (1 load)
            (this.page_details.stage == null)                                   || throwf(new Error("WS course_restore, step 7 state unexpected."));
            const course_id = (this.page_details.mdl_course as DeepPartial<MDL_Course>).id as number;
            await this.page_call({dom_submit: "stage complete submit"});
            await this.page_loaded("page-course-view-", {course: course_id});
        
            /*
            // TODO: Edit course names in first section?  And level? ***
            const new_contents = await ws_call({wsfunction: "core_course_get_contents", courseid: new_course.id
                                                                                        || throwf(new Error("Course ID not found."))});
            const new_s0 = await ws_call({wsfunction: "core_course_get_section_x", sectionid: new_contents[0].id});
            await ws_call({wsfunction: "core_course_update_section_x", section: {id: new_s0.id, summary: new_s0.summary.replace("[Course Name]", name)}});
            */

            // End.
            //this.macro_end();

        }

    }


    export class Index_Rebuild_Macro extends Macro {

        prereq: boolean;
        
        private course_id: number;
        private modules_tab_num: number;
        private last_module_tab_num: number;
    

        public init() {

            this.prereq = false;

            // Check course type
            if (this.page_details.page != "course-view-*")                      { return; }
            const course = this.page_details.mdl_course;
            if (!course || course.format != "onetopic" || !course.id) {  return; }
            this.course_id = course.id;

            // Check editing on
            if (!this.page_details.page_window || !this.page_details.page_window.body_class || !this.page_details.page_window.body_class.match(/\bediting\b/)) {
                console.log("index rebuild pre: editing not on");
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
            {} else {return;}
            this.modules_tab_num = modules_tab_num;
            this.last_module_tab_num = last_module_tab_num;
            
            this.progress_max = this.last_module_tab_num - this.modules_tab_num + 2;
            this.prereq = true;

        }


        
        public async content() {
        
            //this.init();

            // Start
            //this.macro_start();

            // TODO: Don't include hidden tabs or topic headings?
        
            const parser = new DOMParser();
        
            // Get list of modules
            await this.page_load("/course/view.php", {id: this.course_id, section: this.modules_tab_num},
                "page-course-view-", {course: this.course_id},
            );
            if (this.page_details.page != "course-view-*")                      { throw new Error("Unexpected page type."); }
            const modules_list = this.page_details.mdl_course_sections.mdl_course_modules as MDL_Course_Modules[];
        
            (modules_list.length == 1)                                          || throwf(new Error("Expected exactly one resource in Modules tab."));
        
            const modules_index = modules_list[0];
        
            this.tabdata.macro_progress_max = 1 + this.last_module_tab_num - this.modules_tab_num + 2;
        
            let index_html = '<div class="textblock">\n';
        
        // modules_list.shift();
        
            for (let section_num = this.modules_tab_num + 1; section_num <= this.last_module_tab_num; section_num++) {
                //const section_num = section.section                                     || throwf(new Error("Module number not found."));
                await this.page_load("/course/view.php", {id: this.course_id, section: section_num},
                                    "page-course-view-", {course: this.course_id});
                const section_full = this.page_details.mdl_course_sections;
                const section_name = (parser.parseFromString(section_full.summary as string, "text/html").querySelector(".header1")
                                                                                        || throwf(new Error("Module name not found."))
                                    ).textContent                                      || throwf(new Error("Module name content not found."));
        
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
                                    + TabData.escapeHTML((part_name.textContent                 || throwf(new Error("Couldn't get text content."))
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
        
            await this.page_load("/course/mod.php", {sesskey: this.tabdata.page_sesskey, update: modules_index.id},
                                    "page-mod-label-mod", {course: this.course_id});
            await this.page_call({mdl_course_modules: {course: this.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
            mdl_course_module_instance: {intro: index_html}}, dom_submit: true});
            await this.page_loaded( "page-course-view-", {course: this.course_id});
        
            // End.
            //this.macro_end();
        
        } 
    


    }


    export class New_Section_Macro extends Macro {

        prereq:                 boolean;

        private feedback_template_id:   number;
        private course_id:              number;
        private new_section_pos:        number|null;

        new_section: DeepPartial<MDL_Course_Sections> & {
            fullname: string;
            name: string;
        }


        public init() {
            
            console.log("new section pre starting");
            this.prereq = false;
            // Get site details
            // TODO: Also customise image link per site

            //let feedback_template_id: number;
            if (this.tabdata.page_wwwroot == "https://otagopoly-moodle.testing.catlearn.nz" ) {
                this.feedback_template_id = 59;
            } else if (this.tabdata.page_wwwroot == "https://moodle.op.ac.nz") {
                this.feedback_template_id = 59;
            } else                                                                      { return; }
        
            // Check page type
            if (this.page_details.page != "course-view-*") {
                console.log("new section pre: wrong page type");
                return;
            }

            // Check editing on
            if (!this.page_details.page_window || !this.page_details.page_window.body_class || !this.page_details.page_window.body_class.match(/\bediting\b/)) {
                console.log("new section pre: editing not on");
                return;
            }

            // Get course details
            console.log("get course details");
            const course = (this.page_details as page_course_view_data).mdl_course;//(await this.page_call({})).mdl_course;
            if (!course) {console.log("new section pre: couldn't get course details"); return;}
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
            if (modules_tab_num) {  } else                                        {  return; }
            if (this.page_details.mdl_course_sections.section <= last_module_tab_num)
            {} else {return;}
            this.new_section_pos = last_module_tab_num + 1;

            this.progress_max = 17;
            this.prereq = true;
            console.log("new section pre success");
        }
        

        
        public async content() {

            //this.init();



            // Start
            //this.macro_start();

            const name = this.new_section.fullname;
            const short_name = this.new_section.name;

            // Add new tab (1 load)
            await this.page_load(  // TODO: Fix for flexsections?
                "/course/changenumsections.php", {courseid: this.course_id, increase: 1, sesskey: this.tabdata.page_sesskey, insertsection: 0},
                "page-course-view-", {course: this.course_id},
            );
            if (this.page_details.page != "course-view-*")                      { throw new Error("Unexpected page type."); }
            let new_section = this.page_details.mdl_course_sections;

            // Move new tab (1 load)
            console.log("Move new tab")
            await this.page_load(
                "/course/view.php", {id: this.course_id, section: new_section.section, sesskey: this.tabdata.page_sesskey, move: this.new_section_pos - new_section.section
                                                                                    /*|| throwf(new Error("WS course section edit, no amount specified."))*/},
                "page-course-view-", {course: this.course_id},
            );
            new_section.section = this.new_section_pos;     

            // Set new tab details (2 loads)
            await this.page_load(
                "/course/editsection.php", {id: new_section.id, sr: new_section.section,
                                                                                    /*|| throwf(new Error("WS course section edit, no amount specified."))*/},
                                                                                    "page-course-editsection", {course: this.course_id},
                                                                                    );
            await this.page_call({mdl_course_sections: {id: new_section.id, name: short_name, x_options: {level: 1},
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
            await this.page_loaded( "page-course-view-", {course: this.course_id});

            // Add pre-topic message (2 loads)
            await this.page_load("/course/mod.php", {id: this.course_id, sesskey: this.tabdata.page_sesskey, sr: new_section.section, add: "label", section: new_section.section},
                                    "page-mod-label-mod", {course: this.course_id},
                                    );
            await this.page_call({mdl_course_modules: {course: this.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
                mdl_course_module_instance: {intro:
                `<p></p>
        
                <p>After you have worked through all of the above topics, and your facilitator provides you with further information in class,
                you're now ready to demonstrate evidence of what you have learnt in this module.
                Please click on the <strong>Assessments</strong> tab above for further information.</p>`.replace(/^        /gm, "")},
                }, dom_submit: true});
            await this.page_loaded( "page-course-view-", {course: this.course_id});

            // Add blank line (2 loads)
            await this.page_load("/course/mod.php", {id: this.course_id, sesskey: this.tabdata.page_sesskey, sr: new_section.section, add: "label", section: new_section.section},
            "page-mod-label-mod", {course: this.course_id},
            );
            await this.page_call({mdl_course_modules: {course: this.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
            mdl_course_module_instance: {intro: ""}, }, dom_submit: true});
            await this.page_loaded( "page-course-view-", {course: this.course_id});

            // Add feedback topic (2 loads)
            await this.page_load("/course/mod.php", {id: this.course_id, sesskey: this.tabdata.page_sesskey, sr: new_section.section, add: "label", section: new_section.section},
            "page-mod-label-mod", {course: this.course_id},
            );
            await this.page_call({mdl_course_modules: {course: this.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
            mdl_course_module_instance: {intro: 
                `<p></p>
        
                <p><strong>YOUR FEEDBACK</strong></p>
        
                <p>We appreciate your feedback about your experience with working through this module.
                Please click the 'Your feedback' link below if you wish to respond to a five-question survey.
                Thanks!</p>`.replace(/^        /gm, "")},
            }, dom_submit: true});
            await this.page_loaded( "page-course-view-", {course: this.course_id});
        
            // Add feedback activity (2 load)
            await this.page_load("/course/mod.php", {id: this.course_id, sesskey: this.tabdata.page_sesskey, sr: new_section.section, add: "feedback", section: new_section.section},
            "page-mod-feedback-mod", {course: this.course_id},
            );
            await this.page_call({mdl_course_modules: {course: this.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
            mdl_course_module_instance: {name: "Your feedback", intro: 
                `<div class="header2"> <i class="fa fa-bullhorn" aria-hidden="true"></i> FEEDBACK</div>
        
                <div class="textblock">
        
                <p><strong>DESCRIPTION</strong></p>
        
                <p>Please help us improve this learning module by answering five questions about your experience.
                This survey is anonymous.</p>
                </div>`.replace(/^        /gm, "")}, }, dom_submit: true});
            await this.page_loaded( "page-course-view-", {course: this.course_id});
            new_section = (this.page_details as page_course_view_data).mdl_course_sections;
            let feedback_act: DeepPartial<MDL_Course_Modules>|null = null;
                for (const module of new_section.mdl_course_modules) {
                    if (!feedback_act || module.id > feedback_act.id) {
                        feedback_act = module;
                    }
                }

            // Configure Feedback activity (3 loads?)            
            await this.page_load("/mod/feedback/edit.php", {id: feedback_act.id, do_show: "templates"},
                                "page-mod-feedback-edit", {cmid: feedback_act.id});
            await this.page_call({page: "mod-feedback-edit", mdl_course_modules: { mdl_course_module_instance: {mdl_feedback_template_id: this.feedback_template_id}}, dom_submit: true});  // TODO: fix;
            await this.page_loaded("page-mod-feedback-use_templ", {});
            await this.page_call({page: "mod-feedback-use_templ", mdl_course_modules: {}, dom_submit: true});
            await this.page_loaded("page-mod-feedback-edit", {});
            
            // Add footer (2 loads).
            await this.page_load("/course/mod.php", {id: this.course_id, sesskey: this.tabdata.page_sesskey, sr: new_section.section, add: "label", section: new_section.section},
            "page-mod-label-mod", {course: this.course_id},
            );
            await this.page_call({mdl_course_modules: {course: this.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
            mdl_course_module_instance: {intro:
                `<p></p>
        
                <p><span style="font-size: xx-small;">
                Image: <a href="https://stock.tookapic.com/photos/12801" target="_blank">Blooming</a>
                by <a href="https://stock.tookapic.com/pawelkadysz" target="_blank">Paweł Kadysz</a>,
                licensed under <a href="https://creativecommons.org/publicdomain/zero/1.0/deed.en" target="_blank">CC0</a>
                </span></p>`.replace(/^        /gm, "")},
            }, dom_submit: true});
            await this.page_loaded( "page-course-view-", {course: this.course_id});

            // End.
            //this.macro_end();

        }
        


    }


    export class New_Topic_Macro extends Macro {


        prereq:         boolean;

        private course_id:      number;
        private section_num:    number;
        private mod_move_to:    number;
        private topic_first:    boolean;

        new_topic_name: string;


        public init() {
            
            console.log("new topic pre starting");
            this.prereq = false;

            // var doc_details = ws_page_call({wsfunction: "x_doc_get_details"});
            if (this.page_details.page != "course-view-*")                      { throw new Error("Unexpected page type."); }
            const course = this.page_details.mdl_course;
            if (course && course.hasOwnProperty("format") && course.format == "onetopic" && course.id) {  } else { return; }
            this.course_id = course.id;


            // Check editing on
            if (!this.page_details.page_window || !this.page_details.page_window.body_class || !this.page_details.page_window.body_class.match(/\bediting\b/)) {
                console.log("new topic pre: editing not on");
                return;
            }

            //const section_url = await this.page_call({id_act: "* get_element_attribute", selector: "#page-navbar a[href*='section=']", attribute: "href"})
            //                                                                            || throwf(new Error("Section breadcrumb not found."));
            //const section_match = section_url.match(/^(https?:\/\/[a-z\-.]+)\/course\/view.php\?id=(\d+)&section=(\d+)$/)
            //                                                                            || throwf(new Error("Section number not found."));
            //const section_num = parseInt(section_match[3]);
        
            //let section = (await ws_call({wsfunction: "core_course_get_contents", courseid: course.id, options: [{name: "sectionnumber", value: section_num}]}))[0];
            let section = this.page_details.mdl_course_sections;
            //const section_num = section.section;
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

            this.progress_max = 12;
            this.prereq = true;
            console.log("new topic pre success");
        }
        


        public async content() {

            //this.init();



            // Start
            //this.macro_start();
    

            const name = this.new_topic_name;
        
            // Create topic heading (4 loads?)
            await this.page_load("/course/mod.php", {id: this.course_id, sesskey: this.tabdata.page_sesskey, sr: 0 /* TODO: remove? */, add: "label", section: this.section_num},
                "page-mod-label-mod", {course: this.course_id},
            );
            await this.page_call({page: "mod-*-mod", mdl_course_modules: {course: this.course_id, //section: section_num, modname: "label",
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
            await this.page_loaded( "page-course-view-", {course: this.course_id});
            
            //section = (await ws_call({wsfunction: "core_course_get_contents", courseid: this.course_id, options: [{name: "sectionnumber", value: section_num}]}))[0];
            let section: MDL_Course_Sections = this.page_details.mdl_course_sections;

            // Move new module.
            await this.page_load("/course/mod.php", {sesskey: this.tabdata.page_sesskey, sr: section.section, copy: section.mdl_course_modules[section.mdl_course_modules.length-1].id},
            "page-course-view-", {},
            );
            await this.page_load("/course/mod.php", {moveto: this.mod_move_to /*???*/, sesskey: this.tabdata.page_sesskey},
                "page-course-view-", {},
            );

            
            // Create topic end message (4 loads?)
            if (this.topic_first) {
                await this.page_load("/course/mod.php", {id: this.course_id, sesskey: this.tabdata.page_sesskey, sr: 0 /* TODO: remove? */, add: "label", section: this.section_num},
                    "page-mod-label-mod", {course: this.course_id},
                );
                await this.page_call({page: "mod-*-mod", mdl_course_modules: {course: this.course_id, //section: section_num, modname: "label",
                    mdl_course_module_instance: { intro: 
                    `<p></p>
    
                    <p>When you have completed the above activities, and your facilitator provides you with further information,
                    please continue to the next topic below—<strong>[xxxxxxx]</strong>.</p>`
                    },
                    }, dom_submit: true});
                await this.page_loaded( "page-course-view-", {course: this.course_id});
                
                //section = (await ws_call({wsfunction: "core_course_get_contents", courseid: this.course_id, options: [{name: "sectionnumber", value: section_num}]}))[0];
                section = this.page_details.mdl_course_sections;
    
                // Move new module.
                await this.page_load("/course/mod.php", {sesskey: this.tabdata.page_sesskey, sr: section.section, copy: section.mdl_course_modules[section.mdl_course_modules.length-1].id},
                "page-course-view-", {},
                );
                await this.page_load("/course/mod.php", {moveto: this.mod_move_to /*???*/, sesskey: this.tabdata.page_sesskey},
                    "page-course-view-", {},
                );
            } else {
                this.tabdata.page_load_count(4);
            }
        
            // Create space (4 loads?)
            await this.page_load("/course/mod.php", {id: this.course_id, sesskey: this.tabdata.page_sesskey, sr: 0 /* TODO: remove? */, add: "label", section: this.section_num},
                "page-mod-label-mod", {course: this.course_id},
            );
            await this.page_call({page: "mod-*-mod", mdl_course_modules: {course: this.course_id, //section: section_num, modname: "label",
                mdl_course_module_instance: { intro: 
                ""
                },
                }, dom_submit: true});
            await this.page_loaded( "page-course-view-", {course: this.course_id});
            
            //section = (await ws_call({wsfunction: "core_course_get_contents", courseid: this.course_id, options: [{name: "sectionnumber", value: section_num}]}))[0];
            section = this.page_details.mdl_course_sections;

            // Move new module.
            await this.page_load("/course/mod.php", {sesskey: this.tabdata.page_sesskey, sr: section.section, copy: section.mdl_course_modules[section.mdl_course_modules.length-1].id},
            "page-course-view-", {},
            );
            await this.page_load("/course/mod.php", {moveto: this.mod_move_to /*???*/, sesskey: this.tabdata.page_sesskey},
                "page-course-view-", {},
            );



            // End.
            //this.macro_end();

        }



    }


    class Test_Macro extends Macro {

        init() {
            this.progress_max = 10;
            this.prereq = true;
        }


        
        public async content() {

            //this.m_log += "test\n";
            //if (this.tabdata.macro_state != 0) {return}
            //this.tabdata.macro_progress = 0;
            //this.tabdata.macro_progress_max = 10;
            //this.tabdata.macro_state = 1;
            //await page_load("/course/view.php", { id: 7015}, "page-course-view", {"course": 7015});
            //alert(await page_call({}));
            //let page_data: Partial<Page_Data>|null = await this.page_call({});


            //let new_page_data: Partial<Page_Data> = { mdl_course_sections: { summary:
            //    page_data.mdl_course_sections.summary.replace("Hello", "Goodbye") /*replace(/https:\/\/moodle.op.ac.nz\/pluginfile.php\/\d+\/course\/section\/\d+\//,
            //        "https://moodle.op.ac.nz/draftfile.php/807782/user/draft/n/")*/
            //}}

            /*
            for (let count =0; count < 20; count++) {
                await sleep(100);
                //box.value = box.value + ".";
                this.macro_progress++;
                try {
                    await this.popup.update();
                } catch(e) {
                    // Do nothing
                }
            }
            */

            //await this.page_call(new_page_data);
            //await this.page_loaded("page-course-view", {});
            //this.macro_progress = 1;
            /*
            for (let count =0; count < 20; count++) {
                await sleep(100);
                //box.value = box.value + ".";
                this.macro_progress++;
                try {
                    await this.popup.update();
                } catch(e) {
                    // Do nothing
                }
            }
            */
            //   try {
            //    await this.popup.update();
            //} catch(e) {
            //    // Do nothing
            //}

            await this.page_load(
                "/course/index.php", {},
                "page-course-index", {course: 1},
            );

            const site_map = await this.page_call({page: "course-index-category?", dom_expand: true});

            let course_id = 7015;
            let course_context = 911164;

            await this.page_load(
                "/backup/restorefile.php", {contextid: course_context},
                "page-backup-restorefile", {course: course_id},
            );

            const message = await this.page_call({page: "backup-restorefile"});

            await browser.downloads.download({url: message.mdl_course.x_backup_url, saveAs: false});

            //this.tabdata.macro_state = 0;

        }


    }


}