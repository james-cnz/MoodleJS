/*
 * Moodle JS OP Panel Script
 * UI scripts.
 */

namespace MJS {




    export class Popup {


        public tabData:             TabData;

        private macro_uis:          Macro_UI[];

        private status_dom:         HTMLFieldSetElement;
        private progress_bar_dom:   HTMLProgressElement;
        private status_running_dom: HTMLDivElement;
        private status_awaiting_dom: HTMLDivElement;
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
                new New_Section_UI(this),
                new New_Topic_UI(this),
                new Backup_UI(this),
                // new Copy_Grades_UI(this)
            ];

            this.status_dom         = document.querySelector<HTMLFieldSetElement>("fieldset#status")!;
            this.progress_bar_dom   = document.querySelector<HTMLProgressElement>("progress#progress_bar")!;
            this.status_running_dom = document.querySelector<HTMLDivElement>("div#status_running")!;
            this.status_awaiting_dom = document.querySelector<HTMLDivElement>("div#status_awaiting_input")!;
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
                this.status_running_dom.style.display = (this.tabData.macro_state > 0) ? "block" : "none";
                this.status_awaiting_dom.style.display = (this.tabData.macro_state == 3) ? "block" : "none";
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
            const bg_page   = await browser.runtime.getBackgroundPage() as unknown as BackgroundWindow;
            this.tabData    = bg_page.mjs_background.getTabData(tab.id as number);
            this.tabData.popup = this;
            const this_popup: Popup = this;
            this.cancel_button_dom.addEventListener("click", function() { this_popup.onCancel(); });
            this.retry_button_dom.addEventListener("click", function() { this_popup.onRetry(); });
            this.skip_button_dom.addEventListener("click", function() { this_popup.onSkip(); });
            this.reset_button_dom.addEventListener("click", function() { this_popup.onReset(); });
            this.update();
        }


        private onCancel() {
            this.tabData.macro_cancel = true;
        }

        private onRetry() {
            this.tabData.macro_input = "retry";
        }

        private onSkip() {
            this.tabData.macro_input = "skip";
        }

        private onReset() {
            void this.tabData.init();
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
        private new_course_button_dom:      HTMLButtonElement;

        constructor(new_popup: Popup) {
            super(new_popup);
            this.new_course_dom             = document.querySelector<HTMLFieldSetElement>("fieldset#new_course")!;
            this.new_course_name_dom        = document.querySelector<HTMLInputElement>("input#new_course_name")!;
            this.new_course_shortname_dom   = document.querySelector<HTMLInputElement>("input#new_course_shortname")!;
            this.new_course_start_dom       = document.querySelector<HTMLInputElement>("input#new_course_start")!;
            this.new_course_button_dom      = document.querySelector<HTMLButtonElement>("button#new_course_button")!;
            const this_ui = this;
            this.new_course_name_dom.addEventListener("input", function() { this_ui.onInput(); });
            this.new_course_shortname_dom.addEventListener("input", function() { this_ui.onInput(); });
            this.new_course_button_dom.addEventListener("click", function() { this_ui.onClick(); });
        }

        public update() {
            this.new_course_dom.style.display = (this.popup.tabData.macro_state == 0 && this.popup.tabData.macros.new_course.prereq) ? "block" : "none";
        }

        private onInput() {
            this.new_course_button_dom.disabled = !(this.new_course_name_dom.value != "" && this.new_course_shortname_dom.value != "");
        }

        private onClick() {

            (this.popup.tabData.macros.new_course as New_Course_Macro).params = {mdl_course: {
                fullname:   this.new_course_name_dom.value,
                shortname:  this.new_course_shortname_dom.value,
                startdate:  (this.new_course_start_dom.valueAsDate as Date).getTime() / 1000
            }};
            void this.popup.tabData.macros.new_course.run();
        }

    }




    class Index_Rebuild_UI extends Macro_UI {

        private index_rebuild_dom:          HTMLFieldSetElement;
        private index_rebuild_button_dom:   HTMLButtonElement;

        constructor(new_popup: Popup) {
            super(new_popup);
            this.index_rebuild_dom          = document.querySelector<HTMLFieldSetElement>("fieldset#index_rebuild")!;
            this.index_rebuild_button_dom   = document.querySelector<HTMLButtonElement>("button#index_rebuild_button")!;
            const this_ui = this;
            this.index_rebuild_button_dom.addEventListener("click", function() { this_ui.onClick(); });
        }

        public update() {
            this.index_rebuild_dom.style.display = (this.popup.tabData.macro_state == 0 && this.popup.tabData.macros.index_rebuild.prereq) ? "block" : "none";
        }

        private onClick() {
            (this.popup.tabData.macros.index_rebuild as Index_Rebuild_Macro).params = {};
            void this.popup.tabData.macros.index_rebuild.run();
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
            const this_ui = this;
            this.new_section_name_dom.addEventListener("input", function() { this_ui.onInput(); });
            this.new_section_shortname_dom.addEventListener("input", function() { this_ui.onInput(); });
            this.new_section_button_dom.addEventListener("click", function() { this_ui.onClick(); });
        }

        public update() {
            this.new_section_dom.style.display = (this.popup.tabData.macro_state == 0 && this.popup.tabData.macros.new_section.prereq) ? "block" : "none";
        }

        private onInput() {
            this.new_section_button_dom.disabled = !(this.new_section_name_dom.value != "" && this.new_section_shortname_dom.value != "");
        }

        private onClick() {
            (this.popup.tabData.macros.new_section as New_Section_Macro).params = {mdl_course_sections: {
                fullname: this.new_section_name_dom.value,
                name: this.new_section_shortname_dom.value
            }};
            void this.popup.tabData.macros.new_section.run();
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
            const this_ui = this;
            this.new_topic_name_dom.addEventListener("input", function() { this_ui.onInput(); });
            this.new_topic_button_dom.addEventListener("click", function() { this_ui.onClick(); });
        }

        public update() {
            this.new_topic_dom.style.display = (this.popup.tabData.macro_state == 0 && this.popup.tabData.macros.new_topic.prereq) ? "block" : "none";
        }

        private onInput() {
            this.new_topic_button_dom.disabled = !(this.new_topic_name_dom.value != "");
        }

        private onClick() {
            (this.popup.tabData.macros.new_topic as New_Topic_Macro).params = { mdl_course_modules: { fullname: this.new_topic_name_dom.value}};
            void this.popup.tabData.macros.new_topic.run();
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
            const this_ui = this;
            this.backup_button_dom.addEventListener("click", function() { this_ui.onClick(); });
            // this.backup_list_dom.addEventListener("input", function() { this_ui.onInput(); });
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
            (this.popup.tabData.macros.backup as Backup_Macro).params = {mdl_user: {username: document.querySelector<HTMLInputElement>("input#backup_username")!.value,
                                                                                    password_plaintext: document.querySelector<HTMLInputElement>("input#backup_password")!.value}};
            void this.popup.tabData.macros.backup.run();
        }

    }



    class Copy_Grades_UI extends Macro_UI {

        private copy_grades_dom:          HTMLFieldSetElement;
        private copy_grades_button_dom:   HTMLButtonElement;

        constructor(new_popup: Popup) {
            super(new_popup);
            this.copy_grades_dom          = document.querySelector<HTMLFieldSetElement>("fieldset#copy_grades")!;
            this.copy_grades_button_dom   = document.querySelector<HTMLButtonElement>("button#copy_grades_button")!;
            const this_ui = this;
            this.copy_grades_button_dom.addEventListener("click", function() { this_ui.onClick(); });
        }

        public update() {
            this.copy_grades_dom.style.display = (this.popup.tabData.macro_state == 0 && this.popup.tabData.macros.copy_grades.prereq) ? "block" : "none";
        }

        private onClick() {
            (this.popup.tabData.macros.copy_grades as Index_Rebuild_Macro).params = {};
            void this.popup.tabData.macros.copy_grades.run();
        }

    }



    // @ts-ignore
    const mjs_popup: Popup = new Popup();




}
