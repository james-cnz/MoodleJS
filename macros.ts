/**
 * Moodle JS Macro Scripts
 * Implements specific tasks.
 */


// eslint-disable-next-line @typescript-eslint/no-unused-vars
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

        private static check_with_skel(data: unknown, skel: unknown, exclude?: Record<string, boolean>): boolean {
            if (skel === undefined) { return true; }
            else if (typeof(skel) == "function")  {
                throw new Error("Can't compare functions");
            } else if (typeof(skel) == "object" && skel !== null) {
                if (typeof(data) != "object" || data === null) { return false; }
                for (const prop in skel) if (skel.hasOwnProperty(prop) && skel[prop] != undefined) {
                    if (exclude && exclude.hasOwnProperty(prop)) continue;
                    if (!data.hasOwnProperty(prop) || !this.check_with_skel(data[prop], skel[prop])) {
                        throw new Error("Mismatch on property: \"" + prop + "\".  Expected: \"" + skel[prop] + "\", found: \"" + data[prop] + "\".");
                    }
                }
                return true;
            } else {
                return (data == skel);
            }
        }




        public macro_state:         -1|0|1|2|3 = 0;     // -1 error / 0 idle / 1 running / 2 running & awaiting load / 3 paused
        public macro_allow_interrupt: boolean = false;
        // macro_callstack:    string[3]       // 0: macro  1: macro step  2: tabdata function
        public macro_log:           string = "";
        public macro_progress:      number = 1;
        public macro_progress_max:  number = 1;
        public macro_input:         "cancel"|"interrupt"|"retry"|"skip"|null = null;

        public macros: {[index: string] : Macro} = {
            new_course:     new New_Course_Macro(this),
            index_rebuild:  new Index_Rebuild_Macro(this),
            index_rebuild_mt: new Index_Rebuild_MT_Macro(this),
            new_section:    new New_Section_Macro(this),
            new_topic:      new New_Topic_Macro(this),
            backup:         new Backup_Macro(this),
            // copy_grades:    new Copy_Grades_Macro(this)
        };

        public popup:           Popup|null = null;

        private page_last_wwwroot: string;
        // private page_details:   Page_Data|null = null;
        private page_tab_id:    number;

        // private page_load_wait: number = 0;
        private page_message:   Page_Data|Errorlike|null = null;
        private page_is_loaded: boolean = false;


        constructor(tab_id: number) {
            this.page_tab_id = tab_id;
            void this.init();
        }


        public async init() {
            let page_details: Page_Data|null = null;
            try {
                if (this.popup) {
                    this.popup.close();
                }
            } catch (e) { }
            // this.page_load_wait = 0;
            this.page_message = null;
            this.page_is_loaded = true;
            this.macro_log = "";
            this.macro_progress = 1;
            this.macro_progress_max = 1;
            this.macro_input = null;
            this.macro_allow_interrupt = false;
            try {
                this.macro_state = 1;
                page_details = await this.page_call({});
                this.macro_state = 0;
            } catch (e) {
                this.macro_state = 0;
                page_details = null;
            }
            this.macros_init(page_details);
            this.macro_state = 0;
            this.update_ui();
        }


        public update_ui(): void {

            if (this.macro_state == 0) {
                browser.browserAction.setBadgeBackgroundColor({color: "green", tabId: this.page_tab_id});
                browser.browserAction.setBadgeText({text: "", tabId: this.page_tab_id});
            } else if (this.macro_state > 0 && this.macro_state < 3) {
                browser.browserAction.setBadgeBackgroundColor({color: "green", tabId: this.page_tab_id});
                browser.browserAction.setBadgeText({text: ">", tabId: this.page_tab_id});
            } else if (this.macro_state == 3) {
                browser.browserAction.setBadgeBackgroundColor({color: "yellow", tabId: this.page_tab_id});
                browser.browserAction.setBadgeText({text: "||", tabId: this.page_tab_id});
            } else {
                browser.browserAction.setBadgeBackgroundColor({color: "red", tabId: this.page_tab_id});
                browser.browserAction.setBadgeText({text: "X", tabId: this.page_tab_id});
            }

            try {
                if (this.popup) {
                    this.popup.update();
                }
            } catch (e) { }
        }


        public async page_call<T extends Page_Data>(message: DeepPartial<T>): Promise<T> {
            // console.group("page_call");
            // console.debug("page_call start");

            (this.macro_state == 1)                                             || throwf(new Error("Page call:\nUnexpected state."));

            if (message.dom_submit) {
                this.macro_state = 2;
                this.page_is_loaded = false;
                this.page_message = null;
            }
            const result = await browser.tabs.sendMessage(this.page_tab_id, message) as T|Errorlike;
            if (is_Errorlike(result))                                            { throw (new Error(result.message)); }
            this.page_last_wwwroot = result.moodle_page.wwwroot;

            // console.debug("page_call end");
            // console.groupEnd();
            return result;
        }


        public async page_load<T extends Page_Data>(
                                page_data: DeepPartial<T> & {location: {pathname: string, search: {[index: string]: number|string}}},
                                count: number = 1): Promise<T> {
            // console.debug("page_load (calls page_load2)");
            return await this.page_load2<T>(page_data, page_data, count);
        }


        public async page_load2<T extends Page_Data>(
            page_data1: DeepPartial<Page_Data|T> & {location: {pathname: string, search: {[index: string]: number|string}}},
            page_data2: DeepPartial<T>,
            count: number = 1
        ): Promise<T> {
            // console.group("page_load2");
            // console.debug("page_load2 start");
            const pathname = page_data1.location.pathname;
            const search = page_data1.location.search;

            (this.macro_state == 1)                                             || throwf(new Error("Page load:\nUnexpected state."));
            (pathname.match(/(?:\/[a-z]+)+\.php/))                              || throwf(new Error("Page load:\nPathname unexpected."));

            this.macro_state = 2;
            this.page_is_loaded = false;
            this.page_message = null;

            await browser.tabs.update(this.page_tab_id, {url: this.page_last_wwwroot + pathname + "?" + TabData.json_to_search_params(search)});
            // console.debug("page_load2 call page_loaded");
            // console.groupEnd();
            return await this.page_loaded<T>(page_data2, count);
        }


        public async page_loaded<T extends Page_Data>(page_data: DeepPartial<T>,
            count: number = 1): Promise<T> {
            // console.group("page_loaded");
            // console.debug("page_loaded start");

            (this.macro_state == 2)                                             || throwf(new Error("Page loaded:\nUnexpected state."));

            let page_load_wait: number  = 0;
            let page_loaded_time: number|null = null;
            do {
                await sleep(100);   // May lock up here
                page_load_wait += 1;
                if (is_Errorlike(this.page_message))                            {
                    const new_error: Error&{fileName?: string; lineNumber?: number} = new Error(this.page_message.message); // , this.page_message.fileName, this.page_message.lineNumber);
                    new_error.fileName = this.page_message.fileName;
                    new_error.lineNumber = this.page_message.lineNumber;
                    throw new_error;
                }
                if (this.macro_input == "cancel")                               { throw new Error("Cancelled"); }
                if (this.macro_input == "interrupt")                            { throw new Error("Interrupted"); }
                if (this.page_is_loaded && !page_loaded_time) {
                    page_loaded_time = page_load_wait;
                }
                if (page_loaded_time && (page_load_wait - page_loaded_time > 30 * 60 * 10)) { throw new Error("Timed out"); }
                if (page_load_wait <= count * 10) {  // Assume a step takes 1 second
                    this.page_load_count(1 / 10);
                }
            } while (!(this.page_is_loaded && this.page_message));
            this.macro_state = 1;
            const page_details: Page_Data = this.page_message;
            this.page_message = null;
            if (page_load_wait <= count * 10) {
                this.page_load_count(count - page_load_wait / 10);
            }
            if (!this.page_load_match<T>(page_details, page_data))              { throw new Error("Page loaded:\nBody ID or class unexpected."); }
            this.macro_state = 1;
            // console.debug("page_loaded end");
            // console.groupEnd();
            return page_details;
        }


        public page_load_count(count: number = 1): void {
            this.macro_progress += count;
            try {
                if (this.popup) {
                    this.popup.update_progress();
                }
            } catch (e) { }
        }


        public onMessage(message: Page_Data|Errorlike, _sender: browser.runtime.MessageSender) {
            // console.group("onMessage");
            // console.debug("onMessage start");
            let page_details: Page_Data|null = null;

            if (!this.page_message || this.macro_state == 0) {
                // console.debug("expected message");
                this.page_message = message;
                if (is_Errorlike(message)) {
                    page_details = null;
                } else {
                    page_details = message;
                    this.page_last_wwwroot = page_details.moodle_page.wwwroot;
                }

                if (this.page_is_loaded ) {

                    if (this.macro_state == 0) { this.macros_init(page_details); }
                    this.update_ui();
                }
            } else if (!is_Errorlike(this.page_message)) {
                // console.debug("unexpected message");
                this.page_message = new Error("Unexpected page message");
            }

            // console.debug("onMessage end");
            // console.groupEnd();
        }


        public onTabUpdated(_tab_id: number, update_info: Partial<browser.tabs.Tab>, _tab: browser.tabs.Tab): void {
            // console.group("onTabUpdated");
            // console.debug("onTabUpdated start");

            if (update_info && update_info.status) {
                // console.debug("onTabUpdated status: " + update_info.status);

                if (this.macro_state == 0) {

                    if (update_info.status == "loading") {
                        // console.debug("loading");
                        this.page_is_loaded = false;
                        this.page_message = null;
                    } else if (update_info.status == "complete") {
                        // console.debug("complete");
                        this.page_is_loaded = true;
                        this.macros_init(is_Errorlike(this.page_message) ? null : this.page_message);
                    }

                } else if (this.macro_state > 0) {

                    if (!this.page_is_loaded) {
                        if (update_info.status == "complete") {
                            // console.debug("complete");
                            this.page_is_loaded = true;
                        }
                    } else if (!is_Errorlike(this.page_message)) {
                        this.page_message = new Error("Unexpected tab update");
                    }

                }

                this.update_ui();

            }

            // console.debug("onTabUpdated end");
            // console.groupEnd();
        }


        private macros_init(page_details: Page_Data|null) {
            this.macros.new_course.init(page_details);
            this.macros.index_rebuild.init(page_details);
            this.macros.index_rebuild_mt.init(page_details);
            this.macros.new_section.init(page_details);
            this.macros.new_topic.init(page_details);
            this.macros.backup.init(page_details);
            // this.macros.copy_grades.init(page_details);
        }


        private page_load_match<T extends Page_Data_Base>(page_details: Page_Data_Base, page_data: DeepPartial<T>):
            page_details is T {
            // const result = true;
            // if (!page_data.page || page_details.moodle_page.body_id.match(RegExp("^page-" + page_data.page + "$"))) { /* OK */ } else    { result = false; }
            /*for (const prop in body_class) if (body_class.hasOwnProperty(prop)) {
                if ((" " + this.page_details.moodle_page.body_class + " ").match(" " + prop + (body_class[prop] ? ("-" + body_class[prop]) : "") + " "))
                    {  OK  }
                else
                    { result = false; }
            }*/
            return TabData.check_with_skel(page_details/*.moodle_page*/, page_data/*.moodle_page*/, {location: true});
            // return result;
        }


    }




    abstract class Macro {


        public prereq:      boolean     = false;

        public params:      Record<string, unknown>|null     = null;

        protected tabdata:  TabData;

        protected data:     Record<string, unknown>|null     = null;

        protected progress_max: number  = 1;

        protected page_details: Page_Data;

        constructor(new_tabdata: TabData) {
            this.tabdata = new_tabdata;
            this.prereq = false;
        }


        public abstract init(page_details: Page_Data|null): void;


        public async run(): Promise<void> {

            this.params = JSON.parse(JSON.stringify(this.params));

            if (this.tabdata.macro_state != 0) {
                return;
            }

            this.tabdata.macro_input    = null;
            this.tabdata.macro_state    = 1;
            this.tabdata.macro_progress = 0;
            this.tabdata.macro_progress_max = this.progress_max;

            try {
                await this.content();
            } catch (e) {
                // if (e.message != "Cancelled") {
                    this.tabdata.macro_state = -1;
                    // this.tabdata.macro_error = e;
                    this.tabdata.macro_log += /*"Error type:" + this.tabData.macro_error.name + "\n"
                    +*/ e.message + "\n"
                    + ((e.message != "Cancelled" && e.message != "Interrupted" && e.message != "Too many errors" && e.fileName) ? ("file: " + e.fileName + " line: " + e.lineNumber + "\n") : "")
                    + "\n";

                // }
            } // finally {
                this.tabdata.macro_allow_interrupt = false;
                if (this.tabdata.macro_log) {
                    this.tabdata.macro_state = -1;
                    this.tabdata.update_ui();
                    return;
                }
            // }

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


        protected abstract content(): Promise<void>;


    }




    export class Backup_Macro extends Macro {

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

        public params: {mdl_user: {username: string, password_plaintext: string}, exclude_list: number[]}|null = null; // mdl_course_categories: { mdl_course: {id: number}[]} };

        protected login_check_needed: boolean = false;

        public init(page_details: Page_Data) {
            this.prereq     = false;
            if (!page_details || page_details.page != "course-management") { return; }
            this.progress_max = 1000;
            this.page_details = page_details;
            this.data       = {};
            this.prereq     = true;
        }

        protected async login_check() {
            if (this.login_check_needed) {
                const login_start_progress:  number  = this.tabdata.macro_progress;
                try {
                    // alert("before try");
                    this.page_details = await this.tabdata.page_load({location: {pathname: "/my/index.php", search: {}}, page: "my-index"});
                    // alert("after try");
                    this.tabdata.page_load_count(1);
                } catch (e) {
                    // alert("starting catch");
                    // alert(e.message);
                    // alert(e.message != "Unexpected tab update");
                    if (   e.message != "Unexpected tab update"
                        && e.message != "Mismatch on property: \"page\".  Expected: \"my-index\", found: \"login-index\".") throw e;
                    this.tabdata.macro_progress = login_start_progress;
                    // alert("reset stuff");
                    await sleep(4 * 1000);
                    // do_sleep = true;
                    this.tabdata.macro_state = 1;
                    // this.tabdata.update_ui();
                    // alert("check local login");
                    // TODO: pause, reset?
                    // this.page_details = await this.tabdata.page_loaded({page: "local-otago-login"});
                    this.page_details = await this.tabdata.page_call({});
                    if (this.page_details.page == "local-otago-login") {
                        // alert("call click");
                        this.page_details = await this.tabdata.page_call({page: "local-otago-login", dom_submit: "other_users"});
                        // alert("await login");
                        this.page_details = await this.tabdata.page_loaded({page: "login-index"});
                        // alert("call click");
                    } else if (this.page_details.page == "login-index") {
                        this.tabdata.page_load_count(1);
                    } else {
                        throw new Error("Not on login page");
                    }
                    this.page_details = await this.tabdata.page_call({page: "login-index", mdl_user: {username: this.params!.mdl_user.username, password: this.params!.mdl_user.password_plaintext}, dom_submit: "log_in"});
                    // alert("await my");
                    this.page_details = await this.tabdata.page_loaded({page: "my-index"});  // If error here, maybe bad login details?
                    // alert("all OK");
                }
                this.login_check_needed = false;

            } else {
                this.tabdata.page_load_count(2);
            }
        }

        protected exclude_from(course_list: {course_id: number, fullname?: string, shortname?: string}[]) {
            const exclude_list: number[] = this.params!.exclude_list;
            const exclude_set: Set<number> = new Set(exclude_list);
            for (let course_count = 0; course_count < course_list.length; ) {
                if (exclude_set.has(course_list[course_count].course_id)) {
                    course_list.splice(course_count, 1);
                } else {
                    course_count++;
                }
            }
        }

        protected async content() {

            let course_list: {course_id: number, fullname?: string, shortname?: string}[] = [];

            if (false) {

                let tick_change: boolean;
                do {
                    this.page_details = await this.tabdata.page_call<page_course_management_data>({page: "course-management"});
                    const site_map: page_course_management_category = this.page_details.mdl_course_category;
                    tick_change = Backup_Macro.expand_ticked(site_map);
                    if (tick_change) {
                        this.page_details = await this.tabdata.page_call<page_course_management_data>({page: "course-management", mdl_course_category: site_map});
                        // await sleep(1000);
                    }
                } while (tick_change);

                // Get ticked categories.
                // this.page_details = await this.tabdata.page_call<page_course_management_data>({page: "course-management"});
                const category_list = Backup_Macro.ticked_categories(this.page_details.mdl_course_category);

                // Estimate max progress
                let course_count = 0;
                for (const category of category_list) {
                    course_count += category.coursecount;
                }
                this.progress_max = category_list.length + 1 + course_count * 17 + 1;
                this.tabdata.macro_progress_max = this.progress_max;

                // Get course list from category list
                // let course_list: page_course_management_course[] = [];
                for (const category of category_list) {
                    this.page_details = await this.tabdata.page_load<page_course_management_data>({location: {pathname: "/course/management.php", search: {categoryid: category.course_category_id, perpage: 999}}});
                    course_list = course_list.concat(this.page_details.mdl_courses);
                }

                this.exclude_from(course_list);

                // Calculate max progress.
                this.progress_max = category_list.length + 1 + course_list.length * 17 + 1;
                this.tabdata.macro_progress_max = this.progress_max;

            } else {

                // Get ticked categories.
                this.page_details = await this.tabdata.page_call<page_course_management_data>({page: "course-management"});
                const category_list = Backup_Macro.ticked_categories(this.page_details.mdl_course_category);

                // Find course list query.
                this.page_details = await this.tabdata.page_load<page_admin_report_customsql_index_data>({location: {pathname: "/report/customsql/index.php", search: {}}, page: "admin-report-customsql-index"});
                let course_list_query_id: number|null = null;
                for (const query_cat of this.page_details.query_cats) {
                    for (const query of query_cat.mdl_report_customsql_queries) {
                        if (query.displayname.toUpperCase() == "COURSE LIST") { course_list_query_id = query.id; }
                    }
                }
                if (course_list_query_id === null) { throw new Error("Course list query not found"); }

                // Run course list query.
                this.page_details = await this.tabdata.page_load<page_admin_report_customsql_view_data>({location: {pathname: "/report/customsql/view.php", search: {id: course_list_query_id}}, page: "admin-report-customsql-view"});
                let cat_path_col:   number|null = null;
                let course_id_col:  number|null = null;
                let course_name_col: number|null = null;
                let col:            number = 0;
                for (const col_name of this.page_details.query_results!.headers) {
                    switch (col_name.toUpperCase()) {
                        case "CAT PATH":        cat_path_col = col; break;
                        case "COURSE ID":       course_id_col = col; break;
                        case "COURSE SHORT NAME": course_name_col = col; break;
                    }
                    col++;
                }
                if (cat_path_col === null || course_id_col === null || course_name_col === null) {
                    throw new Error("Headers not found");
                }

                // Find courses in ticked categories.
                // const course_list: {course_id: number, shortname: string}[] = [];
                for (const query_row of this.page_details.query_results!.data) {
                    let match: boolean = false;
                    for (const cat of category_list) {
                        if ((query_row[cat_path_col] + "/").search("/" + cat.course_category_id + "/") >= 0) { match = true; }
                    }
                    if (match) { course_list.push({course_id: parseInt(query_row[course_id_col]), shortname: query_row[course_name_col]}); }
                }

                this.exclude_from(course_list);

                // Calculate max progress.
                this.progress_max = 3 + course_list.length * 17 + 1;
                this.tabdata.macro_progress_max = this.progress_max;

            }

            // const error_list: {course_id: number, err: Error}[] = [];
            // let cancelled: boolean = false;

            // let do_sleep:                   boolean = false;
            // let delete_errors: number = 0;

            // Logout.
            const logout_start_progress: number = this.tabdata.macro_progress;
            try {
                this.page_details = await this.tabdata.page_load2({location: {pathname: "/login/logout.php", search: {sesskey: this.page_details.moodle_page.sesskey}}},
                                                                    {location: {pathname: "/local/otago/login.php"}});  // TODO: Page load 3?
            } catch (e) {
                if (e.message != "Unexpected tab update") throw e;
                this.tabdata.macro_progress = logout_start_progress + 1;
                await sleep(4 * 1000);
                this.tabdata.macro_state = 1;
                this.page_details = await this.tabdata.page_call({});
                if (this.page_details.page != "local-otago-login" && this.page_details.page != "my-index")  // Second check in case auto logged in.
                    throw new Error("Not on login page");
                this.tabdata.update_ui();
            }
            this.login_check_needed = true;

            // const unattended_errors_max = 50;
            const unattended_delay = 1.2 * 60 * 60;
            // let unattended_errors = 0;

            this.tabdata.macro_allow_interrupt = true;

            for (const course of course_list) { // this.params.mdl_course_categories.mdl_course) {

                const course_start_progress:  number  = this.tabdata.macro_progress;
                let course_tries:   number  = 0;
                let course_skip:    0|1|2 = 0; // 0: Don't skip  1: Skip backup  2: Skip backup and delete
                let backup_try_error: boolean = false;
                let backup_step:    string = "";

                do {

                    if (backup_try_error) { this.tabdata.macro_progress = course_start_progress; }
                    course_tries++;
                    backup_try_error = false;
                    // let course_deleted = false;
                    let course_context:     number|null = null;
                    let backup_filename:    string|null = null;

                    try {

                        backup_step = "pre-backup login check";
                        await this.login_check();
                    /*catch (e) {
                        this.tabdata.macro_log += "In course: " + course.course_id + ", recovering\n";
                        if (e.message == "Cancelled") { throw e; }
                        // error_list.push({course_id: course.id, err: e});
                        this.tabdata.macro_log += /*"Error type:" + this.tabData.macro_error.name + "\n" */ /*
                        e.message + "\n"
                        + (e.fileName ? ("file: " + e.fileName + " line: " + e.lineNumber + "\n") : "")
                        + "\n";
                        await sleep(4 * 1000);
                        // do_sleep = true;
                        this_try_error = true;
                        this.tabdata.macro_state = 1;
                        // TODO: Rewind progress bar
                        this.tabdata.update_ui();
                    } */

                        backup_step = "creating";

                    // let backup_finished:    boolean     = false;


                    // /*if (!this_try_error && !course_skip)*/ try {

                        // Create backup file (9 loads)
                        this.page_details = await this.tabdata.page_load({location: {pathname: "/backup/backup.php", search: {id: course.course_id}}, page: "backup-backup"});
                        const course_context_match = this.page_details.moodle_page.body_class.match(/(?:^|\s)context-(\d+)(?:\s|$)/)
                                                                                       || throwf(new Error("Backup macro, create backup:\nContext not found."));
                        course_context = parseInt(course_context_match[1]);
                        await sleep(100);

                        this.page_details = await this.tabdata.page_call({page: "backup-backup", dom_submit: "next"});
                        this.page_details = await this.tabdata.page_loaded({page: "backup-backup"});
                        await sleep(100);

                        this.page_details = await this.tabdata.page_call({page: "backup-backup", dom_submit: "next"});
                        this.page_details = await this.tabdata.page_loaded<page_backup_backup_4_data>({page: "backup-backup"});
                        backup_filename = this.page_details.backup.filename;

                        this.page_details = await this.tabdata.page_call({page: "backup-backup", dom_submit: "perform backup"});
                        this.page_details = await this.tabdata.page_loaded<page_backup_backup_last_data>({page: "backup-backup"}, 5);

                        // Wait for final stage.  TODO: Short timeout and continue if syncronous, long timeout and error if asyncronous?
                        while (this.page_details.stage_user != 5) {
                            await sleep(100);
                            if (this.tabdata.macro_input == "cancel")   { throw new Error("Cancelled"); }
                            if (this.tabdata.macro_input == "interrupt") { throw new Error("Interrupted"); }
                            this.page_details = await this.tabdata.page_call<page_backup_backup_last_data>({page: "backup-backup"});
                        }

                        // Check for continue button.  If not present, wait.
                        let reported_missing_cont: boolean = false;
                        for (let waited: number = 0; waited < 15 * 60 * 10 && !this.page_details.cont_button; waited++) {
                            if (waited > 4 * 10 && !reported_missing_cont) {
                                this.tabdata.macro_log += "In course: " + course.course_id + " " + backup_step + " backup: " + backup_filename + "\n";
                                this.tabdata.macro_log += "Continue button not found\n\n";
                                this.tabdata.update_ui();
                                reported_missing_cont = true;
                            }
                            await sleep(100);
                            if (this.tabdata.macro_input == "cancel")   { throw new Error("Cancelled"); }
                            if (this.tabdata.macro_input == "interrupt") { throw new Error("Interrupted"); }
                            this.page_details = await this.tabdata.page_call<page_backup_backup_last_data>({page: "backup-backup"});
                        }

                        /*
                        this.page_details = await this.tabdata.page_call({page: "backup-backup", dom_submit: "continue"});
                        this.page_details = await this.tabdata.page_loaded<page_backup_restorefile_data>({page: "backup-restorefile"});
                        */

                        this.page_details = await this.tabdata.page_load<page_backup_restorefile_data>(
                            {location: {pathname: "/backup/restorefile.php", search: {contextid: course_context}}, // Need session key?
                            page: "backup-restorefile"},
                        );

                    // backup_finished = true;

                    /*} catch (e) {
                        this.tabdata.macro_log += "In course: " + course.course_id + ", creating backup: " + backup_filename + "\n";
                        if (e.message == "Cancelled") { throw e; }
                        // error_list.push({course_id: course.id, err: e});
                        this.tabdata.macro_log += /*"Error type:" + this.tabData.macro_error.name + "\n" */ /*
                        e.message + "\n"
                        + (e.fileName ? ("file: " + e.fileName + " line: " + e.lineNumber + "\n") : "")
                        + "\n";
                        await sleep(4 * 1000);
                        // do_sleep = true;
                        this_try_error = true;
                        if (e.message == "Can not find data record in database table course.") { course_skip = true; }
                        this.tabdata.macro_state = 1;
                        // TODO: Rewind progress bar
                        this.tabdata.update_ui();
                    }*/

                    /*if (!this_try_error && !course_skip) {*/

                        backup_step = "downloading";

                        // let download_tries:         number  = 0;
                        // let download_finished:  boolean     = false;

                        // do {

                                // download_tries++;
                                // this_try_error = false;


                                // Download backup file (1 load?)
                                const backup_index: number = this.page_details.mdl_course.backups.findIndex(function(value) { return value.filename == backup_filename; });
                                const backup_url: string = this.page_details.mdl_course.backups[backup_index].download_url;
                                const backup_download_id = await browser.downloads.download({url: backup_url, saveAs: false});
                                let backup_download_status: browser.downloads.DownloadItem;
                                do {
                                    await sleep(100);
                                    if (this.tabdata.macro_input == "cancel")   { throw new Error("Cancelled"); }
                                    if (this.tabdata.macro_input == "interrupt") { throw new Error("Interrupted"); }
                                    backup_download_status = (await browser.downloads.search({id: backup_download_id}))[0]; // .state;
                                } while (backup_download_status.state == "in_progress" && !backup_download_status.error);
                                if (backup_download_status.state != "complete") { throw new Error("Download error: " + backup_download_status.error); }
                                this.tabdata.page_load_count(1);
                                /* this_try_error = (backup_download_status.state != "complete");
                                if (this_try_error) {
                                    this.tabdata.macro_log += "In course: " + course.course_id + " downloading backup: " + backup_filename + ", try " + download_tries + "\n";
                                    this.tabdata.macro_log += /*"Error type:" + this.tabData.macro_error.name + "\n" */ /*
                                    "Download error: " + backup_download_status.error + "\n";
                                    this.tabdata.update_ui();
                                    if (download_tries >= 3) {
                                        throw new Error("Too many download errors");
                                    }
                                }
                                await sleep(100); */


                                // download_finished = true;



                        // } while (this_try_error);

                        // this.tabdata.page_load_count(1);
                    // }


                    } catch (e) {
                        this.tabdata.macro_log += "In course: " + course.course_id + " " + backup_step + " backup: " + backup_filename + "\n";
                        if (e.message == "Cancelled") { throw e; }
                        // error_list.push({course_id: course.id, err: e});
                        this.tabdata.macro_log += /*"Error type:" + this.tabData.macro_error.name + "\n" */
                        e.message + "\n"
                        + (e.fileName ? ("file: " + e.fileName + " line: " + e.lineNumber + "\n") : "")
                        + "\n";

                        await sleep(4 * 1000);

                        if (e.message.match(/^(?:Can not|Can't) find data record in database table course./)) {
                            course_skip = 2;
                            // course_skip = true;
                            this.tabdata.macro_progress = course_start_progress + 12;
                        } else {
                            // await sleep(100);
                            // do_sleep = true;
                            backup_try_error = true;
                            // unattended_errors++;
                            this.login_check_needed = true;
                            // TODO: Rewind progress bar
                            this.tabdata.macro_progress = course_start_progress;
                            this.tabdata.macro_state = 3;
                            this.tabdata.macro_input = null;
                            this.tabdata.update_ui();
                            const unattended_action =  (   backup_step == "creating" && e.message == "error/error_zip_packing"
                                                        || backup_step == "creating" && e.message == "final_step_cont_dom is null"
                                                        || backup_step == "downloading" && e.message == "Download error: NETWORK_FAILED"
                                                        || backup_step == "creating" && e.message == "Unexpected tab update"
                                                        || backup_step == "creating" && e.message == "error/directory_not_exists"
                                                        || backup_step == "creating" && e.message.startsWith("Exception - Server error")
                                                        || backup_step == "creating" && e.message.startsWith("Exception - cURL error 23: Failed writing body")
                                                        || backup_step == "creating" && e.message.startsWith('Mismatch on property: "page".')  // course page with weird permissions or missing course
                                                        || backup_step == "downloading" && e.message == "this.page_details.mdl_course.backups[backup_index] is undefined"
                                                        || backup_step == "downloading" && e.message == "Download error: CRASH"
                                                        || backup_step == "creating" && e.message == "stage_user_dom is null"
                                                        || backup_step == "creating" && e.message == "backup_dom.querySelector(...) is null"
                                                        || backup_step == "creating" && e.message == "Timed out"
                                                        // || backup_step == "creating" && e.message.startsWith("Exception - HTTP Error") // didn't work on retry. skip? Don't need?
                                                        // downloading "Download error: FILE_FAILED" -- when rclone was disconnected. stop?
                                                       );

                            let unattended_skip = false;
                            let waited = 0;
                            while (!this.tabdata.macro_input) {
                                if (unattended_action && waited >= unattended_delay * 10) {
                                    if (course_tries >= 3) { unattended_skip = true; }
                                    break;
                                }
                                await sleep(100);
                                waited++;
                            }
                            if (this.tabdata.macro_input == "cancel")           { throw new Error("Cancelled"); }
                            else if (this.tabdata.macro_input == "skip") {
                                course_skip = 2;
                                this.tabdata.macro_log += "*** Backup and delete skipped for course: " + course.course_id + " ***\n\n";
                            } else if (unattended_skip) {
                                course_skip = 1;
                                this.tabdata.macro_log += "*** Backup failed for course: " + course.course_id + " ***\n\n";
                            }
                            // if (this.tabdata.macro_input) { unattended_errors = 0; }
                            this.tabdata.macro_input = null;
                            if (course_skip > 0) { this.tabdata.macro_progress = course_start_progress + 12; }
                        }

                        this.tabdata.macro_state = 1;
                        this.tabdata.update_ui();
                    }


                    const delete_start_progress:  number  = this.tabdata.macro_progress;
                    let delete_tries: number = 0;
                    let delete_try_error: boolean = false;
                    if (backup_filename && course_skip < 2) do {
                        delete_tries++;
                        delete_try_error = false;
                        try {

                            backup_step = "pre-delete login check";
                            await this.login_check();

                            backup_step = "deleting";

                            this.page_details = await this.tabdata.page_call({});
                            if (backup_try_error || (delete_tries > 1) || this.page_details.page != "backup-restorefile") {
                                this.page_details = await this.tabdata.page_load<page_backup_restorefile_data>({location: {pathname: "/backup/restorefile.php", search: {contextid: course_context!}}, page: "backup-restorefile"});
                                if (this.page_details.mdl_course.backups.length == 0) {
                                    throw new Error("Backup file not found: " + backup_filename);
                                }
                            } else {
                                this.tabdata.page_load_count(1);
                            }

                            // Delete backup file (2 loads?)
                            this.page_details = await this.tabdata.page_call({page: "backup-restorefile", dom_submit: "manage"});
                            this.page_details = await this.tabdata.page_loaded({page: "backup-backupfilesedit"});   // *** LOCKS UP HERE IF FILE LIST EMPTY?

                            this.page_details = await this.tabdata.page_call({page: "backup-backupfilesedit", mdl_course: { backups: [{filename: backup_filename, click: true}]}});

                            this.page_details = await this.tabdata.page_call({page: "backup-backupfilesedit", backup: {click: "delete"}});
                            this.page_details = await this.tabdata.page_call({page: "backup-backupfilesedit", backup: {click: "delete_ok"}});
                            this.page_details = await this.tabdata.page_call({page: "backup-backupfilesedit", dom_submit: "save"});
                            this.page_details = await this.tabdata.page_loaded({page: "backup-restorefile"});

                        } catch (e) {
                            this.tabdata.macro_log += "In course: " + course.course_id + " " + backup_step + " backup: " + backup_filename + "\n";
                            if (e.message == "Cancelled") { throw e; }
                            // delete_errors++;

                            // error_list.push({course_id: course.id, err: e});
                            this.tabdata.macro_log += /*"Error type:" + this.tabData.macro_error.name + "\n" */
                            e.message + "\n"
                            + (e.fileName ? ("file: " + e.fileName + " line: " + e.lineNumber + "\n") : "")
                            + "\n";

                            await sleep(4 * 1000);

                            /*if (e.message.match(/^(?:Can not|Can't) find data record in database table context.$/)) {
                                course_skip = true;
                                // course_skip = true;
                                this.tabdata.macro_progress = delete_start_progress + 5;
                            } else*/ if (backup_step == "deleting" && e.message == "Backup file not found: " + backup_filename && (backup_try_error || (delete_tries > 1))) {
                                this.tabdata.macro_progress = delete_start_progress + 5;
                            } else {
                                delete_try_error = true;
                                // unattended_errors++;
                                this.login_check_needed = true;
                                // await sleep(100);
                                // do_sleep = true;

                                // TODO: Rewind progress bar
                                this.tabdata.macro_progress = delete_start_progress;
                                this.tabdata.macro_state = 3;
                                this.tabdata.macro_input = null;
                                this.tabdata.update_ui();
                                const unattended_retry = (backup_step == "deleting" && e.message == "Unexpected tab update"
                                                       || backup_step == "deleting" && e.message == "Timed out"
                                                       || backup_step == "deleting" && e.message == "backup_filemanager_dom is null"
                                                       || backup_step == "deleting" && e.message.startsWith('Mismatch on property: "page".')) // Bad gateway
                                                        && (delete_tries < 3);
                                let waited = 0;
                                while (!this.tabdata.macro_input && !(unattended_retry && waited >= unattended_delay * 10)) {
                                    await sleep(100);
                                    waited++;
                                }
                                if (this.tabdata.macro_input == "cancel")       { throw new Error("Cancelled"); }
                                else if (this.tabdata.macro_input == "skip") {
                                    course_skip = 2;
                                    this.tabdata.macro_log += "*** Delete skipped for course: " + course.course_id + " ***\n\n";
                                }
                                // if (this.tabdata.macro_input) { unattended_errors = 0; }
                                this.tabdata.macro_input = null;
                                if (course_skip >= 2) { this.tabdata.macro_progress = delete_start_progress + 5; }
                            }

                            this.tabdata.macro_state = 1;
                            this.tabdata.update_ui();
                        }
                    } while (delete_try_error && course_skip < 2);
                    else if (course_skip >= 2) { this.tabdata.macro_progress = delete_start_progress + 5; }

                    /* if (course_try_error && !course_skip || delete_errors > 3) {
                        this.tabdata.macro_state = 3;
                        this.tabdata.macro_input = null;
                        this.tabdata.update_ui();
                        while (!this.tabdata.macro_input && !this.tabdata.macro_cancel) {
                            await sleep(100);
                        }
                        if (this.tabdata.macro_cancel)                      { throw new Error("Cancelled"); }
                        course_skip = (this.tabdata.macro_input == "skip");
                        this.tabdata.macro_input = null;
                        this.tabdata.macro_state = 1;
                        this.tabdata.update_ui();
                        course_tries = 0;
                        delete_errors = 0;
                    } */

                    // TODO: If skip, fast forward progress bar
                    // if (course_skip) { this.tabdata.page_load_count(16); }

                } while (backup_try_error && course_skip < 1);

                /*if (this_try_error && !course_skip) {
                    throw new Error("Too many consecutive errors.");
                }*/
            }

            this.tabdata.macro_allow_interrupt = false;


        }


    }





    export type New_Course_Params = {
        mdl_course: { fullname: string, shortname: string, startdate: number, format: "onetopic"|"multitopic"};
    };

    type New_Course_Data = {
        // mdl_course: { template_id: number };
        mdl_course_category: { course_category_id: number; name: string };
    };

    export class New_Course_Macro extends Macro {


        public params: New_Course_Params|null = null;
        protected data: New_Course_Data|null = null;


        public init(page_details: Page_Data) {

            this.prereq = false;

            this.page_details = page_details;

            if (!this.page_details || !this.page_details.hasOwnProperty("page") || this.page_details.page != "course-index(-category)?")     { return; }

            if (!this.page_details.mdl_course_category.course_category_id) { return; }

            this.data = {
                mdl_course_category: this.page_details.mdl_course_category,
                // mdl_course: { template_id: template_id}
            };



            this.progress_max = 17 + 1;
            this.prereq = true;

        }


        protected async content() {  // TODO: Set properties.

            if (!this.data || !this.params)                                     throw new Error("New course macro, prereq:\ndata not set.");

            let template_id: number;
            if (this.page_details.moodle_page.wwwroot == "https://otagopoly-moodle.testing.catlearn.nz" ) {
                template_id = this.params!.mdl_course.format == "onetopic" ? 8310 : 8705;
            } else if (this.page_details.moodle_page.wwwroot == "https://moodle.op.ac.nz") {
                template_id = this.params!.mdl_course.format == "onetopic" ? 8310 : 8705;
            } else if (this.page_details.moodle_page.wwwroot == "http://localhost" || this.page_details.moodle_page.wwwroot == "https://localhost") {
                template_id = this.params!.mdl_course.format == "onetopic" ? 2 : 3;
            } else                                                                      { throw new Error("Site not recognised."); }

            // Get template course context (1 load)
            this.page_details = await this.tabdata.page_load(
                {location: {pathname: "/course/view.php", search: {id: template_id, section: 0}},
                page: "course-view-[a-z]+", mdl_course: {course_id: template_id}}
            );
            const source_context_match = this.page_details.moodle_page.body_class.match(/(?:^|\s)context-(\d+)(?:\s|$)/)!;
            const source_context = parseInt(source_context_match[1]);

            // Load course restore page (1 load)
            this.page_details = await this.tabdata.page_load<page_backup_restorefile_data>(
                {location: {pathname: "/backup/restorefile.php", search: {contextid: source_context}},
                page: "backup-restorefile" /*, mdl_course: {id: this.data.mdl_course.template_id}*/},
            );

            // Click restore backup file (1 load)
            this.page_details = await this.tabdata.page_call<page_backup_restorefile_data>({page: "backup-restorefile", dom_submit: "restore"});
            this.page_details = await this.tabdata.page_loaded<page_backup_restore_data_2>({page: "backup-restore", stage: 2 /*, mdl_course: {template_id: this.data.mdl_course.template_id}*/});

            // Confirm (1 load)
            this.page_details = await this.tabdata.page_call<page_backup_restore_data_2>({page: "backup-restore", stage: 2, dom_submit: "stage 2 submit"});
            this.page_details = await this.tabdata.page_loaded<page_backup_restore_data_4d>({page: "backup-restore", stage: 4, displayed_stage: "Destination" /*, mdl_course: {template_id: this.data.mdl_course.template_id}*/});

            // Destination: Search for category (1 load)
            this.page_details = await this.tabdata.page_call<page_backup_restore_data_4d>({page: "backup-restore", stage: 4, displayed_stage: "Destination", mdl_course_category: {name: this.data.mdl_course_category.name}, dom_submit: "stage 4 new cat search"});
            this.page_details = await this.tabdata.page_loaded<page_backup_restore_data_4d>({page: "backup-restore", stage: 4, displayed_stage: "Destination"});

            // Destination: Select category (1 load)
            this.page_details = await this.tabdata.page_call<page_backup_restore_data_4d>({page: "backup-restore", stage: 4, displayed_stage: "Destination", mdl_course_category: {course_category_id: this.data.mdl_course_category.course_category_id}, dom_submit: "stage 4 new continue"});
            this.page_details = await this.tabdata.page_loaded<page_backup_restore_data_4s>({page: "backup-restore", stage: 4, displayed_stage: "Settings" /*, mdl_course: {template_id: this.data.mdl_course.template_id}*/});
            await sleep(100);

            // Restore settings (1 load)
            if (this.page_details.restore_settings.users != undefined) {
                this.page_details = await this.tabdata.page_call<page_backup_restore_data_4s>({page: "backup-restore", stage: 4, displayed_stage: "Settings", restore_settings: {users: false}});
            }
            this.page_details = await this.tabdata.page_call<page_backup_restore_data_4s>({page: "backup-restore", stage: 4, displayed_stage: "Settings", dom_submit: "stage 4 settings submit"});
            this.page_details = await this.tabdata.page_loaded<page_backup_restore_data_8>({page: "backup-restore", stage: 8 /*, mdl_course: {template_id: this.data.mdl_course.template_id}*/});
            await sleep(100);

            // Course settings (1 load)
            const course = {fullname: this.params.mdl_course.fullname, shortname: this.params.mdl_course.shortname, startdate: this.params.mdl_course.startdate};
            this.page_details = await this.tabdata.page_call<page_backup_restore_data_8>({page: "backup-restore", stage: 8, mdl_course: course, dom_submit: "stage 8 submit"});
            this.page_details = await this.tabdata.page_loaded<page_backup_restore_data_16>({page: "backup-restore", stage: 16 /*, mdl_course: {template_id: this.data.mdl_course.template_id}*/});

            // Review & Process (~5 loads)
            this.page_details = await this.tabdata.page_call<page_backup_restore_data_16>({page: "backup-restore", stage: 16, dom_submit: "stage 16 submit"});
            this.page_details = await this.tabdata.page_loaded<page_backup_restore_data_final>({page: "backup-restore", stage: null /*, mdl_course: {template_id: this.data.mdl_course.template_id}*/}, 5);

            // Wait for final stage.
            while (this.page_details.stage_user != 7) {
                await sleep(100);
                if (this.tabdata.macro_input == "cancel")   { throw new Error("Cancelled"); }
                this.page_details = await this.tabdata.page_call<page_backup_restore_data_final>({page: "backup-restore", stage: null});
            }

            // Wait for continue button.
            let course_id: number|undefined;
            do {
                await sleep(100);
                this.page_details = await this.tabdata.page_call<page_backup_restore_data_final>({page: "backup-restore", stage: null});
                course_id = this.page_details.mdl_course.course_id;
            } while (!course_id);

            // Complete--Go to new course (1 load)
            this.page_details = await this.tabdata.page_call({page: "backup-restore", stage: null, dom_submit: "stage complete submit"});
            this.page_details = await this.tabdata.page_loaded<page_course_view_data>({page: "course-view-[a-z]+", mdl_course: {course_id: course_id}});

            // Turn editing on (1 load).
            if (!this.page_details.moodle_page.body_class.match(/\bediting\b/)) {
                this.page_details = await this.tabdata.page_load<page_course_view_data>(
                    {location: {pathname: "/course/view.php", search: {id: course_id, section: 0, sesskey: this.page_details.moodle_page.sesskey, edit: "on"}},
                    page: "course-view-[a-z]+", mdl_course: {course_id: course_id}},
                );
            } else {
                this.tabdata.page_load_count(1);
            }
            (this.page_details.moodle_page.body_class.match(/\bediting\b/))     || throwf(new Error("New course macro, turn editing on:\nEditing not on."));

            // Fill in course name (2 loads)
            const section_0_id = this.page_details.mdl_course.mdl_course_sections[0].course_section_id || throwf(new Error("New course macro, fill in course name:\nID not found."));
            this.page_details = await this.tabdata.page_load<page_course_editsection_data>(
                {location: {pathname: "/course/editsection.php", search: {id: section_0_id}}, // NOTE: Needs editing on.
                page: "course-editsection" /*, mdl_course: {id: course_id}*/},
            );
            const desc = this.page_details.mdl_course_section.summary.replace(/\[Course Name\]/g, this.params.mdl_course.fullname);
            this.page_details = await this.tabdata.page_call({page: "course-editsection", mdl_course_section: {summary: desc}, dom_submit: true});
            this.page_details = await this.tabdata.page_loaded({page: "course-view-[a-z]+", mdl_course: {course_id: course_id}});

        }


    }




    export class Index_Rebuild_Macro extends Macro {


        protected data: {mdl_course: {course_id: number}; mdl_course_sections: {section: number}; last_section_num: number} | null = null;


        public init(page_details: Page_Data) {

            this.prereq = false;

            this.page_details = page_details;

            // Check course type
            if (!this.page_details || this.page_details.page != "course-view-[a-z]+")                      { return; }
            const course = this.page_details.mdl_course;
            if (!course || course.format != "onetopic" || !course.course_id) {  return; }

            // Check editing on
            if (!this.page_details.moodle_page || !this.page_details.moodle_page.body_class || !this.page_details.moodle_page.body_class.match(/\bediting\b/)) {
                return;
            }

            // Find Modules tab number
            const course_contents = course.mdl_course_sections;
            let modules_tab_num: number|undefined|null = null;
            let last_module_tab_num: number|undefined|null = null;
            for (const section of course_contents) {
                if ((section.options!.level! <= 0) && (section.section <= this.page_details.mdl_course_section!.section) && (section.name.toUpperCase().includes("MODULES"))) {
                    modules_tab_num = section.section;
                    last_module_tab_num = modules_tab_num;
                } else if (last_module_tab_num && section.options!.level! > 0) { // TODO: Need to scrape level property.
                    last_module_tab_num = section.section;
                }
            }
            if (modules_tab_num && last_module_tab_num) {  } else                                        {  return; }
            if (this.page_details.mdl_course_section!.section <= last_module_tab_num)
            { } else { return; }

            this.data = {mdl_course: {course_id: course.course_id}, mdl_course_sections: {section: modules_tab_num}, last_section_num: last_module_tab_num};

            this.progress_max = last_module_tab_num - modules_tab_num + 3 + 1;
            this.prereq = true;

        }


        protected async content() {

            if (!this.data /*|| !this.params*/)                                     throw new Error("Index rebuild macro, prereq:\ndata not set.");
            // TODO: Don't include hidden tabs or topic headings?

            const parser = new DOMParser();

            // Get list of sections (1 load)
            this.page_details = await this.tabdata.page_load<page_course_view_data>(
                {location: {pathname: "/course/view.php", search: {id: this.data.mdl_course.course_id, section: this.data.mdl_course_sections.section}},
                page: "course-view-[a-z]+", mdl_course: {course_id: this.data.mdl_course.course_id}},
            );
            const modules_list = this.page_details.mdl_course_section!.mdl_course_modules!;
            (modules_list.length == 1)                                          || throwf(new Error("Index rebuild macro, get list:\nExpected exactly one resource in Modules tab."));
            const modules_index = modules_list[0];


            // Get section contents (1 load per section)
            let index_html = '<div class="textblock">\n';
            // modules_list.shift();
            for (let section_num = this.data.mdl_course_sections.section + 1; section_num <= this.data.last_section_num; section_num++) {
                // const section_num = section.section                                     || throwf(new Error("Module number not found."));
                this.page_details = await this.tabdata.page_load<page_course_view_data>(
                    {location: {pathname: "/course/view.php", search: {id: this.data.mdl_course.course_id, section: section_num}},
                    page: "course-view-[a-z]+", mdl_course: {course_id: this.data.mdl_course.course_id}});
                const section_full = this.page_details.mdl_course_section!;
                const section_name = (parser.parseFromString(section_full.summary as string, "text/html").querySelector(".header1")!
                                    ).textContent!;
                index_html = index_html
                            + '<a href="' + this.page_details.moodle_page.wwwroot + "/course/view.php?id=" + this.data.mdl_course.course_id + "&section=" + section_num + '"><b>' + TabData.escapeHTML(section_name.trim()) + "</b></a>\n"
                            + "<ul>\n";
                for (const mod of section_full.mdl_course_modules!) {
                    // parse description
                    const mod_desc = parser.parseFromString(mod.intro || "", "text/html");
                    const part_name = mod_desc.querySelector(".header2, .header2gradient")!;
                    if (part_name) {
                        index_html = index_html
                                    + "<li>"
                                    + TabData.escapeHTML((part_name.textContent!).trim())
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
                {location: {pathname: "/course/mod.php", search: {sesskey: this.page_details.moodle_page.sesskey, update: modules_index.course_module_id}}},
                {page: "mod-[a-z]+-mod" /*, mdl_course: {id: this.data.mdl_course.id}*/});
            this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_module: {course: this.data.mdl_course.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
                intro: index_html}, dom_submit: true});
            this.page_details = await this.tabdata.page_loaded({page: "course-view-[a-z]+", mdl_course: {course_id: this.data.mdl_course.course_id}});

        }


    }


    export class Index_Rebuild_MT_Macro extends Macro {


        protected data: {mdl_course: {course_id: number}; mdl_course_sections: {id: number}[]; } | null = null;


        public init(page_details: Page_Data) {

            this.prereq = false;

            this.page_details = page_details;

            // Check course type
            if (!this.page_details || this.page_details.page != "course-view-[a-z]+")                      { return; }
            const course = this.page_details.mdl_course;
            if (!course || course.format != "multitopic" || !course.course_id) {  return; }

            // Check editing on
            if (!this.page_details.moodle_page || !this.page_details.moodle_page.body_class || !this.page_details.moodle_page.body_class.match(/\bediting\b/)) {
                return;
            }

            // Find Modules tab number
            const course_contents = course.mdl_course_sections;
            let sections: {id: number}[] = [];
            let modules_tab_num: number|undefined|null = null;
            let last_module_tab_num: number|undefined|null = null;
            for (const section of course_contents) {
                if (section.options && (section.options!.level! <= 0) && (section.section <= this.page_details.mdl_course_section!.section) && (section.name.toUpperCase().includes("MODULES"))) {
                    modules_tab_num = section.section;
                    sections = [];
                    sections.push({id: section.course_section_id!});
                    last_module_tab_num = modules_tab_num;
                } else if (last_module_tab_num && section.options && section.options!.level! == 1) { // TODO: Need to scrape level property.
                    sections.push({id: section.course_section_id!});
                    last_module_tab_num = section.section;
                }
            }
            if (modules_tab_num && last_module_tab_num) {  } else                                        {  return; }
            if (this.page_details.mdl_course_section!.section <= last_module_tab_num)
            { } else { return; }

            this.data = {mdl_course: {course_id: course.course_id}, mdl_course_sections: sections};

            this.progress_max = sections.length + 2 + 1;
            this.prereq = true;

        }


        protected async content() {

            if (!this.data /*|| !this.params*/)                                     throw new Error("Index rebuild macro, prereq:\ndata not set.");
            // TODO: Don't include hidden tabs or topic headings?

            // const parser = new DOMParser();

            // Get list of sections (1 load)
            this.page_details = await this.tabdata.page_load<page_course_view_data>(
                {location: {pathname: "/course/view.php", search: {id: this.data.mdl_course.course_id, sectionid: this.data.mdl_course_sections[0].id}},
                page: "course-view-[a-z]+", mdl_course: {course_id: this.data.mdl_course.course_id}},
            );
            const modules_list = this.page_details.mdl_course_section!.mdl_course_modules!;
            (modules_list.length == 1)                                          || throwf(new Error("Index rebuild macro, get list:\nExpected exactly one resource in Modules tab."));
            const modules_index = modules_list[0];


            // Get section contents (1 load per section)
            let index_html = '<div class="textblock">\n';
            // modules_list.shift();
            for (const section of this.data.mdl_course_sections) {
                if (section == this.data.mdl_course_sections[0]) { continue; }
                // const section_num = section.section                                     || throwf(new Error("Module number not found."));
                this.page_details = await this.tabdata.page_load<page_course_view_data>(
                    {location: {pathname: "/course/view.php", search: {id: this.data.mdl_course.course_id, sectionid: section.id}},
                    page: "course-view-[a-z]+", mdl_course: {course_id: this.data.mdl_course.course_id}});
                const section_full = this.page_details.mdl_course_section!;
                const section_name = section_full.name.trim();
                index_html = index_html
                            + '<a href="' + this.page_details.moodle_page.wwwroot + "/course/view.php?id=" + this.data.mdl_course.course_id + "&sectionid=" + section.id + '"><b>' + TabData.escapeHTML(section_name.trim()) + "</b></a>\n"
                            + "<ul>\n";
                let section_count = 0;
                while (this.page_details.mdl_course.mdl_course_sections[section_count].course_section_id! != section_full.course_section_id!) {
                    section_count++;
                }
                section_count++;
                let subsection: page_course_view_course_section;
                while (this.page_details.mdl_course.mdl_course_sections.length > section_count
                        && (subsection = this.page_details.mdl_course.mdl_course_sections[section_count]).options
                        && subsection.options!.level == 2) {
                /* for (const mod of section_full.mdl_course_modules!) { */
                    // parse description
                    // const mod_desc = parser.parseFromString(mod.intro || "", "text/html");
                    // const part_name = subsection.name; // mod_desc.querySelector(".header2, .header2gradient")!;
                    // if (part_name) {
                        index_html = index_html
                                    + "<li>"
                                    + TabData.escapeHTML(subsection.name.trim())
                                    + "</li>\n";
                    // }
                    section_count++;
                }
                index_html = index_html
                            + "</ul>\n"
                            + "<br />\n";
            }
            index_html = index_html
                        + "</div>\n";

            // Update TOC (2 loads)
            this.page_details = await this.tabdata.page_load2(
                {location: {pathname: "/course/mod.php", search: {sesskey: this.page_details.moodle_page.sesskey, update: modules_index.course_module_id}}},
                {page: "mod-[a-z]+-mod" /*, mdl_course: {id: this.data.mdl_course.id}*/});
            this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_module: {course: this.data.mdl_course.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
                intro: index_html}, dom_submit: true});
            this.page_details = await this.tabdata.page_loaded({page: "course-view-[a-z]+", mdl_course: {course_id: this.data.mdl_course.course_id}});

        }


    }


    export class New_Section_Macro extends Macro {


        public params: {mdl_course_sections: {name: string, fullname: string}}|null = null;
        protected data: {mdl_course: {course_id: number}, mdl_course_section: {section: number}, mdl_course_module: {feedback_template_id: number}}|null = null;


        public init(page_details: Page_Data) {

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
            } else if (this.page_details.moodle_page.wwwroot == "http://localhost" || this.page_details.moodle_page.wwwroot == "https://localhost") {
                feedback_template_id = 1;
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
                if (section.options!.level! <= 0 && section.section <= this.page_details.mdl_course_section!.section && section.name.toUpperCase().includes("MODULES")) {
                    modules_tab_num = section.section;
                    last_module_tab_num = modules_tab_num;
                } else if (last_module_tab_num && section.options!.level! > 0) { // TODO: Need to scrape level property.
                    last_module_tab_num = section.section;
                }
            }
            if (modules_tab_num && last_module_tab_num) {  } else                                        { return; }
            if (this.page_details.mdl_course_section!.section <= last_module_tab_num)
            { } else { return; }
            // this.new_section_pos = last_module_tab_num + 1;

            this.data = {mdl_course: {course_id: course.course_id}, mdl_course_section: {section: last_module_tab_num + 1}, mdl_course_module: {feedback_template_id: feedback_template_id}};

            this.progress_max = 4 + 1;
            this.prereq = true;
        }


        protected async content() {

            if (!this.data || !this.params)                                     throw new Error("New section macro, prereq:\ndata not set.");

            // Add new tab (1 load)
            this.page_details = await this.tabdata.page_load2<page_course_view_data>(  // TODO: Fix for flexsections?
                {location: {pathname: "/course/changenumsections.php", search: {courseid: this.data.mdl_course.course_id, increase: 1, sesskey: this.page_details.moodle_page.sesskey, insertsection: 0}}},
                {page: "course-view-[a-z]+", mdl_course: {course_id: this.data.mdl_course.course_id}},
            );
            const new_section = this.page_details.mdl_course_section!;

            // Move new tab (1 load)
            this.page_details = await this.tabdata.page_load(
                {location: {pathname: "/course/view.php", search: {id: this.data.mdl_course.course_id, section: new_section.section, sesskey: this.page_details.moodle_page.sesskey, move: this.data.mdl_course_section.section - new_section.section},
                                                                                    /*|| throwf(new Error("WS course section edit, no amount specified."))*/},
                page: "course-view-[a-z]+", mdl_course: {course_id: this.data.mdl_course.course_id}},
            );
            new_section.section = this.data.mdl_course_section.section;

            // Set new tab details (2 loads)
            this.page_details = await this.tabdata.page_load(
                {location: {pathname: "/course/editsection.php", search: {id: new_section.course_section_id!, sr: new_section.section,
                                                                                    /*|| throwf(new Error("WS course section edit, no amount specified."))*/}},
                                                                                    page: "course-editsection", /*mdl_course: {id: this.data.mdl_course.id}*/},
                                                                                    );
            this.page_details = await this.tabdata.page_call({page: "course-editsection", mdl_course_section: {course_section_id: new_section.course_section_id, name: this.params.mdl_course_sections.name, options: {level: 1},
                summary:
                `<div class="header1"> <i class="fa fa-list" aria-hidden="true"></i> ${this.params.mdl_course_sections.fullname}</div>

                <p></p>

                <p>[Module intro]</p>

                <p>You can tick the boxes down the right-hand side of the screen to track your progress through this module.
                Boxes with a dashed border will check themselves when you complete an activity.</p>`.replace(/^        /gm, ""),
            }, dom_submit: true});
            this.page_details = await this.tabdata.page_loaded({page: "course-view-[a-z]+", mdl_course: {course_id: this.data.mdl_course.course_id}});

            // Add post-topic message (2 loads)
            // Body ID or class unexpected.
            // this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {id: this.data.mdl_course.course_id, sesskey: this.page_details.moodle_page.sesskey, sr: new_section.section, add: "label", section: new_section.section}}},
            //                         {page: "mod-[a-z]+-mod", /*mdl_course: {id: this.data.mdl_course.id}*/},
            //                         );
            // this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_module: {course: this.data.mdl_course.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
            //     intro:
            //     `<p></p>

            //     <p>After you have worked through all of the above topics, and your facilitator provides you with further information in class,
            //     you're now ready to demonstrate evidence of what you have learnt in this module.
            //     Please click on the <strong>Assessments</strong> tab above for further information.</p>`.replace(/^        /gm, "")},
            //     dom_submit: true});
            // this.page_details = await this.tabdata.page_loaded({page: "course-view-[a-z]+", mdl_course: {course_id: this.data.mdl_course.course_id}});

            // Add blank line (2 loads)
            // this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {id: this.data.mdl_course.course_id, sesskey: this.page_details.moodle_page.sesskey, sr: new_section.section, add: "label", section: new_section.section}}},
            // {page: "mod-[a-z]+-mod" /*, mdl_course: {id: this.data.mdl_course.id}*/},
            // );
            // this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_module: {course: this.data.mdl_course.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
            // intro: "", }, dom_submit: true});
            // this.page_details = await this.tabdata.page_loaded({page: "course-view-[a-z]+" /*, mdl_course: {id: this.data.mdl_course.id}*/});

            // Add feedback topic (2 loads)
            // this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {id: this.data.mdl_course.course_id, sesskey: this.page_details.moodle_page.sesskey, sr: new_section.section, add: "label", section: new_section.section}}},
            //     {page: "mod-[a-z]+-mod" /*, mdl_course: {id: this.data.mdl_course.id}*/},
            // );
            // this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_module: {course: this.data.mdl_course.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
            // intro:
            //     `<p></p>

            //     <p><strong>YOUR FEEDBACK</strong></p>

            //     <p>We appreciate your feedback about your experience with working through this module.
            //     Please click the 'Your feedback' link below if you wish to respond to a five-question survey.
            //     Thanks!</p>`.replace(/^        /gm, "")},
            // dom_submit: true});
            // this.page_details = await this.tabdata.page_loaded({page: "course-view-[a-z]+", mdl_course: {course_id: this.data.mdl_course.course_id}});

            // Add feedback activity (2 load)
            // this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {id: this.data.mdl_course.course_id, sesskey: this.page_details.moodle_page.sesskey, sr: new_section.section, add: "feedback", section: new_section.section}}},
            //     {page: "mod-[a-z]+-mod" /*, mdl_course: {id: this.data.mdl_course.id}*/},
            // );
            // this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_module: {course: this.data.mdl_course.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
            // name: "Your feedback", intro:
            //     `<div class="header2"> <i class="fa fa-bullhorn" aria-hidden="true"></i> FEEDBACK</div>

            //     <div class="textblock">

            //     <p><strong>DESCRIPTION</strong></p>

            //     <p>Please help us improve this learning module by answering five questions about your experience.
            //     This survey is anonymous.</p>
            //     </div>`.replace(/^        /gm, ""), }, dom_submit: true});
            // this.page_details = await this.tabdata.page_loaded<page_course_view_data>({page: "course-view-[a-z]+", mdl_course: {course_id: this.data.mdl_course.course_id}});
            // new_section = this.page_details.mdl_course_section!;
            // let feedback_act: page_course_view_course_module|null = null;
            //     for (const module of new_section.mdl_course_modules) {
            //         if (!feedback_act || module.course_module_id > feedback_act.course_module_id) {
            //             feedback_act = module;
            //         }
            //     }

            // Configure Feedback activity (3 loads?)
            // this.page_details = await this.tabdata.page_load({location: {pathname: "/mod/feedback/edit.php", search: {id: feedback_act.course_module_id, do_show: "templates"}},
            //                     page: "mod-feedback-edit", /*mdl_course_modules: {id: feedback_act.id}*/});
            // this.page_details = await this.tabdata.page_call({page: "mod-feedback-edit", mdl_course_module: { mdl_feedback_template_id: this.data.mdl_course_module.feedback_template_id}, dom_submit: true});  // TODO: fix;
            // this.page_details = await this.tabdata.page_loaded({page: "mod-feedback-use_templ"});
            // this.page_details = await this.tabdata.page_call({page: "mod-feedback-use_templ", dom_submit: true});
            // this.page_details = await this.tabdata.page_loaded({page: "mod-feedback-edit"});

            // Add footer (2 loads).
            // this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {id: this.data.mdl_course.course_id, sesskey: this.page_details.moodle_page.sesskey, sr: new_section.section, add: "label", section: new_section.section}}},
            // {page: "mod-[a-z]+-mod" /*, mdl_course: {id: this.data.mdl_course.id}*/},
            // );
            // this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_module: {course: this.data.mdl_course.course_id, /*sectionnum: new_section.section,*/ /*modname: "label",*/
            // intro:
            //     `<p></p>

            //     <p><span style="font-size: xx-small;">
            //     Image: <a href="https://stock.tookapic.com/photos/12801" target="_blank">Blooming</a>
            //     by <a href="https://stock.tookapic.com/pawelkadysz" target="_blank">Paweł Kadysz</a>,
            //     licensed under <a href="https://creativecommons.org/publicdomain/zero/1.0/deed.en" target="_blank">CC0</a>
            //     </span></p>`.replace(/^        /gm, ""),
            // }, dom_submit: true});
            // this.page_details = await this.tabdata.page_loaded({page: "course-view-[a-z]+", mdl_course: {course_id: this.data.mdl_course.course_id}});

        }


    }




    export class New_Topic_Macro extends Macro {


        public params: {mdl_course_modules: {fullname: string}}|null = null;

        protected data: {mdl_course: {course_id: number}, mdl_course_section: {section: number}, mdl_course_module: {moveto?: number, first_topic: boolean}}|null = null;

        public init(page_details: Page_Data) {

            this.prereq = false;

            this.page_details = page_details;

            if (!this.page_details || this.page_details.page != "course-view-[a-z]+")                      { return; }
            const course = this.page_details.mdl_course;
            if (course && course.hasOwnProperty("format") && course.format == "onetopic" && course.course_id) {  } else { return; }


            // Check editing on
            if (!this.page_details.moodle_page || !this.page_details.moodle_page.body_class || !this.page_details.moodle_page.body_class.match(/\bediting\b/)) {
                return;
            }

            // Find Modules tab number
            const course_contents = course.mdl_course_sections;
            let modules_tab_num: number|undefined|null = null;
            let last_module_tab_num: number|undefined|null = null;
            for (const section of course_contents) {
                if ((section.options!.level! <= 0) && (section.section <= this.page_details.mdl_course_section!.section) && (section.name.toUpperCase().includes("MODULES"))) {
                    modules_tab_num = section.section;
                    last_module_tab_num = modules_tab_num;
                } else if (last_module_tab_num && section.options!.level! > 0) { // TODO: Need to scrape level property.
                    last_module_tab_num = section.section;
                }
            }
            if (modules_tab_num && last_module_tab_num) {  } else                                        {  return; }
            if (this.page_details.mdl_course_section!.section <= last_module_tab_num)
            { } else { return; }
            // this.new_section_pos = last_module_tab_num + 1;

            const section = this.page_details.mdl_course_section!;

            if (section.section == modules_tab_num) { return; }

            let mod_pos = section.mdl_course_modules!.length - 1;
            let mod_match_pos = 3;

            while (mod_pos > -1 && mod_match_pos > -1) {
                if (mod_match_pos == 3 && section.mdl_course_modules![mod_pos].mdl_module_name == "label" && section.mdl_course_modules![mod_pos].name.toUpperCase().match(/\bIMAGE\b/)) {
                    mod_match_pos -= 1;
                    mod_pos -= 1;
                } else if (mod_match_pos == 3) {
                    mod_match_pos -= 1;
                } else if (mod_match_pos == 2 && section.mdl_course_modules![mod_pos].mdl_module_name == "feedback" && section.mdl_course_modules![mod_pos].name.toUpperCase().match(/\bFEEDBACK\b/)) {
                    mod_match_pos -= 1;
                    mod_pos -= 1;
                } else if (mod_match_pos == 1 && section.mdl_course_modules![mod_pos].mdl_module_name == "label" && section.mdl_course_modules![mod_pos].name.toUpperCase().match(/\bFEEDBACK\b/)) {
                    mod_match_pos -= 1;
                    mod_pos -= 1;
                } else if (mod_match_pos == 0 && section.mdl_course_modules![mod_pos].mdl_module_name == "label" && section.mdl_course_modules![mod_pos].name == "") {
                    mod_pos -= 1;
                } else if (mod_match_pos == 0 && section.mdl_course_modules![mod_pos].mdl_module_name == "label" && section.mdl_course_modules![mod_pos].name.replace(/\s+/g, " ").toUpperCase().match(/\bASSESSMENTS TAB\b/)) {
                    mod_match_pos -= 1;
                    mod_pos -= 1;
                } else {
                    break;
                }
            }

            const topic_first = (mod_pos < 0) ? true : false;

            this.data = {mdl_course: {course_id: course.course_id}, mdl_course_section: {section: section.section}, mdl_course_module: {
                moveto: (mod_pos + 1 < section.mdl_course_modules!.length) ? (section.mdl_course_modules![mod_pos + 1].course_module_id) : undefined,
                first_topic: topic_first}};

            this.progress_max = 8 + 1;
            this.prereq = true;
        }


        protected async content() {

            if (!this.data || !this.params)                                     throw new Error("New section macro, prereq:\ndata not set.");

            const name = this.params.mdl_course_modules.fullname;

            if (!this.data.mdl_course_module.first_topic) {

                // Create space (4 loads?)
                this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {id: this.data.mdl_course.course_id, sesskey: this.page_details.moodle_page.sesskey, add: "label", section: this.data.mdl_course_section.section}}},
                    {page: "mod-[a-z]+-mod" /*, mdl_course: {id: this.data.mdl_course.id}*/},
                );
                this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_module: {course: this.data.mdl_course.course_id, // section: section_num, modname: "label",
                    intro:
                    ""
                    }, dom_submit: true});
                this.page_details = await this.tabdata.page_loaded<page_course_view_data>({page: "course-view-[a-z]+", mdl_course: {course_id: this.data.mdl_course.course_id}});

                const section = this.page_details.mdl_course_section!;

                // Move new module.
                if (this.data.mdl_course_module.moveto) {
                    this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {sesskey: this.page_details.moodle_page.sesskey, sr: section.section, copy: section.mdl_course_modules![section.mdl_course_modules!.length - 1].course_module_id}}},
                    {page: "course-view-[a-z]+"},
                    );
                    this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {moveto: this.data.mdl_course_module.moveto, sesskey: this.page_details.moodle_page.sesskey}}},
                        {page: "course-view-[a-z]+"},
                    );
                } else {
                    this.tabdata.page_load_count(2);
                }

            } else {
                this.tabdata.page_load_count(4);
            }

            // Create topic heading (4 loads?)
            this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {id: this.data.mdl_course.course_id, sesskey: this.page_details.moodle_page.sesskey, add: "label", section: this.data.mdl_course_section.section}}},
                {page: "mod-[a-z]+-mod" /*, mdl_course: {id: this.data.mdl_course.id}*/},
            );
            this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_module: {course: this.data.mdl_course.course_id, // section: section_num, modname: "label",
            intro:
                `<div class="header2"> <i class="fa fa-align-justify" aria-hidden="true"></i> ${name}</div>

                <div class="textblock">

                <p>[Topic introduction, including learning objectives]</p>

                <p><strong>INSTRUCTIONS</strong></p>

                <p>[Topic instructions]</p>

                </div>`.replace(/^        /gm, ""),
            }, dom_submit: true});
            this.page_details = await this.tabdata.page_loaded<page_course_view_data>({page: "course-view-[a-z]+", mdl_course: {course_id: this.data.mdl_course.course_id}});


            // Move new module.
            if (this.data.mdl_course_module.moveto) {
                const section = this.page_details.mdl_course_section!;
                this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {sesskey: this.page_details.moodle_page.sesskey, sr: section.section, copy: section.mdl_course_modules![section.mdl_course_modules!.length - 1].course_module_id}}},
                    {page: "course-view-[a-z]+"},
                );
                this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {moveto: this.data.mdl_course_module.moveto, sesskey: this.page_details.moodle_page.sesskey}}},
                    {page: "course-view-[a-z]+"},
                );
            } else {
                this.tabdata.page_load_count(2);
            }

            // Create topic end message (4 loads?)
            // if (this.data.mdl_course_module.first_topic) {
            //     this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {id: this.data.mdl_course.course_id, sesskey: this.page_details.moodle_page.sesskey, add: "label", section: this.data.mdl_course_section.section}}},
            //         {page: "mod-[a-z]+-mod" /*, mdl_course: {id: this.data.mdl_course.id}*/},
            //     );
            //     this.page_details = await this.tabdata.page_call({page: "mod-[a-z]+-mod", mdl_course_module: {course: this.data.mdl_course.course_id, // section: section_num, modname: "label",
            //         intro:
            //         `<p></p>

            //         <p>When you have completed the above activities, and your facilitator provides you with further information,
            //         please continue to the next topic below—<strong>[xxxxxxx]</strong>.</p>`
            //         }, dom_submit: true});
            //     this.page_details = await this.tabdata.page_loaded<page_course_view_data>({page: "course-view-[a-z]+", mdl_course: {course_id: this.data.mdl_course.course_id}});

            //     section = this.page_details.mdl_course_section!;

            //     // Move new module.
            //     this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {sesskey: this.page_details.moodle_page.sesskey, sr: section.section, copy: section.mdl_course_modules[section.mdl_course_modules.length - 1].course_module_id}}},
            //         {page: "course-view-[a-z]+"},
            //     );
            //     this.page_details = await this.tabdata.page_load2({location: {pathname: "/course/mod.php", search: {moveto: this.data.mdl_course_module.next_id /*???*/, sesskey: this.page_details.moodle_page.sesskey}}},
            //         {page: "course-view-[a-z]+"},
            //     );
            // } else {
            //     this.tabdata.page_load_count(4);
            // }
        }


    }



    export type Copy_Grades_Params = Record<string, never>;

    type Copy_Grades_Data = {
        grades_table_as_text: string;
    };

    export class Copy_Grades_Macro extends Macro {


        public params: Copy_Grades_Params|null = null;
        protected data: Copy_Grades_Data|null = null;


        public init(page_details: Page_Data) {

            this.prereq = false;
            this.page_details = page_details;
            if (!this.page_details || this.page_details.page != "grade-report-grader-index") { return; }
            this.data = { grades_table_as_text: (page_details as page_grade_report_grader_index_data).grades_table_as_text };
            this.prereq = true;

        }

        protected async content() {
            await navigator.clipboard.writeText(this.data!.grades_table_as_text);
        }

    }




}
