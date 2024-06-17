/*
 * Moodle JS OP Panel Script
 * UI scripts.
 */


import "./browser_polyfill_mv3.js";
import {DeepPartial} from "./shared.js";
import {TabData} from "./macros.js"


// eslint-disable-next-line @typescript-eslint/no-unused-vars


    export type popup_input = "init2"|"update"|"update_progress"|"close"|null;


    export class Popup {


        public tabPort:             browser.runtime.Port;
        public tabData:             TabData;

        private macro_uis:          Macro_UI[];

        private status_dom:         HTMLFieldSetElement;
        private progress_bar_dom:   HTMLProgressElement;
        private status_running_dom: HTMLDivElement;
        private status_paused_dom:  HTMLDivElement;
        private interrupt_allowed_dom: HTMLDivElement;
        private interrupt_button_dom:  HTMLButtonElement;
        private cancel_button_dom:  HTMLButtonElement;
        private retry_button_dom:   HTMLButtonElement;
        private skip_button_dom:    HTMLButtonElement;
        private status_error_dom:   HTMLDivElement;
        private error_message_dom:  HTMLTextAreaElement;
        private reset_button_dom:   HTMLButtonElement;


        constructor() {

            this.macro_uis = [
                new New_Course_UI(this),
                new Index_Rebuild_UI(this),
                new Index_Rebuild_MT_UI(this),
                new New_Section_UI(this),
                new New_Topic_UI(this),
                new Backup_UI(this),
                // new Copy_Grades_UI(this)
            ];

            this.status_dom         = document.querySelector<HTMLFieldSetElement>("fieldset#status")!;
            this.progress_bar_dom   = document.querySelector<HTMLProgressElement>("progress#progress_bar")!;
            this.status_running_dom = document.querySelector<HTMLDivElement>("div#status_running")!;
            this.status_paused_dom  = document.querySelector<HTMLDivElement>("div#status_paused")!;
            this.interrupt_allowed_dom = document.querySelector<HTMLDivElement>("div#interrupt_allowed")!;
            this.interrupt_button_dom  = document.querySelector<HTMLButtonElement>("button#interrupt_button")!;
            this.cancel_button_dom  = document.querySelector<HTMLButtonElement>("button#cancel_button")!;
            this.retry_button_dom   = document.querySelector<HTMLButtonElement>("button#retry_button")!;
            this.skip_button_dom    = document.querySelector<HTMLButtonElement>("button#skip_button")!;
            this.status_error_dom   = document.querySelector<HTMLDivElement>("div#status_error")!;
            this.error_message_dom  = document.querySelector<HTMLTextAreaElement>("textarea#error_message")!;
            this.reset_button_dom   = document.querySelector<HTMLButtonElement>("button#reset_button")!;

            void this.init();

        }


        public update() {

            for (const macro_ui of this.macro_uis) {
                macro_ui.update();
            }

            this.status_dom.style.display = (this.tabData.macro_state == 0) ? "none" : "block";
            if (this.tabData.macro_state != 0) {
                this.progress_bar_dom.value = this.tabData.macro_progress;
                this.progress_bar_dom.max   = this.tabData.macro_progress_max;
                this.status_running_dom.style.display   = (this.tabData.macro_state > 0) ? "block" : "none";
                this.interrupt_allowed_dom.style.display = (this.tabData.macro_allow_interrupt && this.tabData.macro_state != 3) ? "block" : "none";
                this.status_paused_dom.style.display    = (this.tabData.macro_state == 3) ? "block" : "none";
                // if (this.tabData.macro_state < 0) {
                    this.error_message_dom.value = this.tabData.macro_log; /*"Error type:" + this.tabData.macro_error.name + "\n"
                                                +*/ // this.tabData.macro_error!.message + "\n"
                                                // + (this.tabData.macro_error!.fileName ? ("file: " + this.tabData.macro_error!.fileName + " line: " + this.tabData.macro_error!.lineNumber + "\n") : "");
                // }
                this.status_error_dom.style.display = (this.tabData.macro_state < 0 || this.tabData.macro_log) ? "block" : "none";
                this.reset_button_dom.style.display = (this.tabData.macro_state < 0) ? "inline-block" : "none";
            }

        }


        public update_progress() {
            this.progress_bar_dom.value = this.tabData.macro_progress;
        }


        public close() {
            window.close();
        }


        private async init() {
            const tab       = (await browser.tabs.query({active: true, currentWindow: true}))[0];
            this.tabPort = browser.runtime.connect({ name: "popup " + tab.id!.toString() });
            this.tabPort.onMessage.addListener((message) => { this.onBGMessage(message as TabData); });
        }


        public init2() {
            this.cancel_button_dom.addEventListener("click", () => { this.onCancel(); });
            this.interrupt_button_dom.addEventListener("click", () => { this.onInterrupt(); });
            this.retry_button_dom.addEventListener("click", () => { this.onRetry(); });
            this.skip_button_dom.addEventListener("click", () => { this.onSkip(); });
            this.reset_button_dom.addEventListener("click", () => { this.onReset(); });
            this.update();
        }


        private onBGMessage(message: TabData) {
            this.tabData = message;
            switch (message.popup_input) {
                case "init2": this.init2(); break;
                case "update": this.update(); break;
                case "update_progress": this.update_progress(); break;
                case "close": this.close(); break;
            }
        }


        public postBGMessage(message: DeepPartial<TabData>) {
            this.tabPort.postMessage(message);
        }


        private onCancel() {
            this.postBGMessage({macro_input: "cancel"});
        }

        private onInterrupt() {
            this.postBGMessage({macro_input: "interrupt"});
        }

        private onRetry() {
            this.postBGMessage({macro_input: "retry"});
        }

        private onSkip() {
            this.postBGMessage({macro_input: "skip"});
        }

        private onReset() {
            this.postBGMessage({macro_input: "init"});
            this.close();
        }


    }




    abstract class Macro_UI {

        protected popup: Popup;

        constructor(new_popup: Popup) {
            this.popup = new_popup;
        }

        public abstract update(): void;

    }




    class New_Course_UI extends Macro_UI {

        private new_course_dom:             HTMLFieldSetElement;
        private new_course_name_dom:        HTMLInputElement;
        private new_course_shortname_dom:   HTMLInputElement;
        private new_course_start_dom:       HTMLInputElement;
        private new_course_format_dom:      NodeListOf<HTMLInputElement>;
        private new_course_button_dom:      HTMLButtonElement;

        constructor(new_popup: Popup) {
            super(new_popup);
            this.new_course_dom             = document.querySelector<HTMLFieldSetElement>("fieldset#new_course")!;
            this.new_course_name_dom        = document.querySelector<HTMLInputElement>("input#new_course_name")!;
            this.new_course_shortname_dom   = document.querySelector<HTMLInputElement>("input#new_course_shortname")!;
            this.new_course_start_dom       = document.querySelector<HTMLInputElement>("input#new_course_start")!;
            this.new_course_format_dom      = document.querySelectorAll<HTMLInputElement>("input[name='new_course_format']");
            this.new_course_button_dom      = document.querySelector<HTMLButtonElement>("button#new_course_button")!;
            this.new_course_name_dom.addEventListener("input", () => { this.onInput(); });
            this.new_course_shortname_dom.addEventListener("input", () => { this.onInput(); });
            this.new_course_format_dom[0].addEventListener("click", () => { this.onInput(); });
            this.new_course_format_dom[1].addEventListener("click", () => { this.onInput(); });
            this.new_course_button_dom.addEventListener("click", () => { this.onClick(); });
        }

        public update() {
            this.new_course_dom.style.display = (this.popup.tabData.macro_state == 0 && this.popup.tabData.macros.new_course.prereq) ? "block" : "none";
        }

        private onInput() {
            this.new_course_button_dom.disabled = !(this.new_course_name_dom.value != "" && this.new_course_shortname_dom.value != ""
                                                    && (this.new_course_format_dom[0].checked || this.new_course_format_dom[1].checked));
        }

        private onClick() {
            this.popup.postBGMessage({
                macro_input: "run",
                macros: { new_course: { params: {
                    mdl_course: {
                        fullname:   this.new_course_name_dom.value,
                        shortname:  this.new_course_shortname_dom.value,
                        startdate:  (this.new_course_start_dom.valueAsDate!).getTime() / 1000,
                        format:     this.new_course_format_dom[0].checked ? "onetopic" : "multitopic"
                    }
                }}}
            });
        }

    }




    class Index_Rebuild_UI extends Macro_UI {

        private index_rebuild_dom:          HTMLFieldSetElement;
        private index_rebuild_button_dom:   HTMLButtonElement;

        constructor(new_popup: Popup) {
            super(new_popup);
            this.index_rebuild_dom          = document.querySelector<HTMLFieldSetElement>("fieldset#index_rebuild")!;
            this.index_rebuild_button_dom   = document.querySelector<HTMLButtonElement>("button#index_rebuild_button")!;
            this.index_rebuild_button_dom.addEventListener("click", () => { this.onClick(); });
        }

        public update() {
            this.index_rebuild_dom.style.display = (this.popup.tabData.macro_state == 0 && this.popup.tabData.macros.index_rebuild.prereq) ? "block" : "none";
        }

        private onClick() {
            this.popup.postBGMessage({
                macro_input: "run",
                macros: { index_rebuild: { params: {} } }
            });
        }

    }


    class Index_Rebuild_MT_UI extends Macro_UI {

        private index_rebuild_mt_dom:       HTMLFieldSetElement;
        private index_rebuild_mt_button_dom: HTMLButtonElement;

        constructor(new_popup: Popup) {
            super(new_popup);
            this.index_rebuild_mt_dom       = document.querySelector<HTMLFieldSetElement>("fieldset#index_rebuild_mt")!;
            this.index_rebuild_mt_button_dom = document.querySelector<HTMLButtonElement>("button#index_rebuild_mt_button")!;
            this.index_rebuild_mt_button_dom.addEventListener("click", () => { this.onClick(); });
        }

        public update() {
            this.index_rebuild_mt_dom.style.display = (this.popup.tabData.macro_state == 0 && this.popup.tabData.macros.index_rebuild_mt.prereq) ? "block" : "none";
        }

        private onClick() {
            this.popup.postBGMessage({
                macro_input:"run",
                macros: { index_rebuild_mt: { params: {} } }
            });
        }

    }


    class New_Section_UI extends Macro_UI {

        private new_section_dom:            HTMLFieldSetElement;
        private new_section_name_dom:       HTMLInputElement;
        private new_section_shortname_dom:  HTMLInputElement;
        private new_section_button_dom:     HTMLButtonElement;

        constructor(new_popup: Popup) {
            super(new_popup);
            this.new_section_dom            = document.querySelector<HTMLFieldSetElement>("fieldset#new_section")!;
            this.new_section_name_dom       = document.querySelector<HTMLInputElement>("input#new_section_name")!;
            this.new_section_shortname_dom  = document.querySelector<HTMLInputElement>("input#new_section_shortname")!;
            this.new_section_button_dom     = document.querySelector<HTMLButtonElement>("button#new_section_button")!;
            this.new_section_name_dom.addEventListener("input", () => { this.onInput(); });
            this.new_section_shortname_dom.addEventListener("input", () => { this.onInput(); });
            this.new_section_button_dom.addEventListener("click", () => { this.onClick(); });
        }

        public update() {
            this.new_section_dom.style.display = (this.popup.tabData.macro_state == 0 && this.popup.tabData.macros.new_section.prereq) ? "block" : "none";
        }

        private onInput() {
            this.new_section_button_dom.disabled = !(this.new_section_name_dom.value != "" && this.new_section_shortname_dom.value != "");
        }

        private onClick() {
            this.popup.postBGMessage({
                macro_input: "run",
                macros: { new_section: { params: {
                    mdl_course_sections: {
                        fullname: this.new_section_name_dom.value,
                        name: this.new_section_shortname_dom.value
                    }
                }}}
            });
        }

    }



    class New_Topic_UI extends Macro_UI {

        private new_topic_dom:          HTMLFieldSetElement;
        private new_topic_name_dom:     HTMLInputElement;
        private new_topic_button_dom:   HTMLButtonElement;

        constructor(new_popup: Popup) {
            super(new_popup);
            this.new_topic_dom          = document.querySelector<HTMLFieldSetElement>("fieldset#new_topic")!;
            this.new_topic_name_dom     = document.querySelector<HTMLInputElement>("input#new_topic_name")!;
            this.new_topic_button_dom   = document.querySelector<HTMLButtonElement>("button#new_topic_button")!;
            this.new_topic_name_dom.addEventListener("input", () => { this.onInput(); });
            this.new_topic_button_dom.addEventListener("click", () => { this.onClick(); });
        }

        public update() {
            this.new_topic_dom.style.display = (this.popup.tabData.macro_state == 0 && this.popup.tabData.macros.new_topic.prereq) ? "block" : "none";
        }

        private onInput() {
            this.new_topic_button_dom.disabled = !(this.new_topic_name_dom.value != "");
        }

        private onClick() {
            this.popup.postBGMessage({
                macro_input: "run",
                macros: { new_topic: { params: {
                    mdl_course_modules: { fullname: this.new_topic_name_dom.value}
                }}}
            });
        }

    }


    class Backup_UI extends Macro_UI {

        private backup_dom: HTMLFieldSetElement;
        // private backup_list_dom: HTMLTextAreaElement;
        private backup_button_dom: HTMLButtonElement;
        // private backup_params: { mdl_course_categories: { mdl_course: {id: number}[]} }|null = null;

        constructor(new_popup: Popup) {
            super(new_popup);
            this.backup_dom         = document.querySelector<HTMLFieldSetElement>("fieldset#backup")!;
            // this.backup_list_dom    = document.querySelector<HTMLTextAreaElement>("textarea#backup_list")!;
            this.backup_button_dom  = document.querySelector<HTMLButtonElement>("button#backup_button")!;
            this.backup_button_dom.addEventListener("click", () => { this.onClick(); });
            // this.backup_list_dom.addEventListener("input", () => { this.onInput(); });
        }

        public update() {
            this.backup_dom.style.display = (this.popup.tabData.macro_state == 0 && this.popup.tabData.macros.backup.prereq) ? "block" : "none";
        }

        /*
        private onInput() {

            try {
                const backup_list: string[] = this.backup_list_dom.value.trim().split(/[\r\n]+/);

                // Read headings
                let id_col: number|null = null;
                {
                    const backup_headings_txt: string = backup_list.shift() + ",";
                    const line_regexp = /\s*("[^"]*"|'[^']*'|[^,'"]*)\s*,/y;
                    let col: number = 0;
                    while (line_regexp.lastIndex < backup_headings_txt.length) {
                        const backup_heading = line_regexp.exec(backup_headings_txt)![1];
                        if (backup_heading.trim().match(/^['"]?id['"]?$/) && id_col === null) {
                            id_col = col;
                        }
                        col++;
                    }
                    if (id_col === null) { throw new Error("ID not found"); }
                }

                // Read rows
                this.backup_params = { mdl_course_categories: { mdl_course: []}};

                for (const backup_row_unterminated of backup_list) {
                    const backup_row_txt = backup_row_unterminated + ",";
                    const line_regexp = /\s*("[^"]*"|'[^']*'|[^,'"]*)\s*,/y;
                    let col: number = 0;
                    while (col < id_col) {
                        line_regexp.exec(backup_row_txt)![1];
                        col++;
                    }
                    const backup_cell = line_regexp.exec(backup_row_txt)![1];
                    const id = parseInt(backup_cell.trim().match(/^['"]?([0-9]+)['"]?$/)![1]);
                    this.backup_params.mdl_course_categories.mdl_course.push({id: id});
                }
                this.backup_button_dom.disabled = false;

            } catch (e) {
                this.backup_params = null;
                this.backup_button_dom.disabled = true;
            }

        }
        */

        private onClick() {
            const exclude_input: string = document.querySelector<HTMLTextAreaElement>("textarea#backup_exclude_list")!.value;
            const exclude_strings: string[] = exclude_input.split(/\r?\n/);
            const exclude_list: number[] = [];
            for (const exclude_string of exclude_strings) {
                const exclude_match_1 = exclude_string.match(/^(\d+)$/);
                if (exclude_match_1) { exclude_list.push(parseInt(exclude_match_1[1])); }
                const exclude_match_2 = exclude_string.match(/^backup-moodle2-course-(\d+)-\S+-\d{8}-\d{4}(?:-nu)?.mbz$/);
                if (exclude_match_2) { exclude_list.push(parseInt(exclude_match_2[1])); }
            }
            this.popup.postBGMessage({
                macro_input: "run",
                macros: { backup: { params: {
                    mdl_user: {
                        username: document.querySelector<HTMLInputElement>("input#backup_username")!.value,
                        password_plaintext: document.querySelector<HTMLInputElement>("input#backup_password")!.value
                    },
                    include_users: document.querySelector<HTMLInputElement>("input#backup_include_users")!.checked,
                    exclude_list: exclude_list
                }}}
            });
        }

    }



    // @ts-expect-error: Declared but not used
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    class Copy_Grades_UI extends Macro_UI {

        private copy_grades_dom:          HTMLFieldSetElement;
        private copy_grades_button_dom:   HTMLButtonElement;

        constructor(new_popup: Popup) {
            super(new_popup);
            this.copy_grades_dom          = document.querySelector<HTMLFieldSetElement>("fieldset#copy_grades")!;
            this.copy_grades_button_dom   = document.querySelector<HTMLButtonElement>("button#copy_grades_button")!;
            this.copy_grades_button_dom.addEventListener("click", () => { this.onClick(); });
        }

        public update() {
            this.copy_grades_dom.style.display = (this.popup.tabData.macro_state == 0 && this.popup.tabData.macros.copy_grades.prereq) ? "block" : "none";
        }

        private onClick() {
            this.popup.postBGMessage({
                macro_input: "run",
                macros: { copy_grades: { params: {} } }
            });
        }

    }



    // @ts-expect-error: Declared but value not read
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    const mjs_popup: Popup = new Popup();




