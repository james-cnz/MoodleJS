/**
 * Moodle JS Background Script
 * Stores data for each tab, and passes on messages from content scripts.
 */


namespace MJS {

    export class Background {

        private tabData:    { [index: number]: TabData } = { };

        public constructor() {
            const bg_this = this;
            browser.runtime.onMessage.addListener(
                function(message: object, sender: browser.runtime.MessageSender) {
                    bg_this.onMessage(message as Page_Data_Out, sender);
                }
            );
            browser.tabs.onUpdated.addListener(
                function(tab_id: number, update_info: Partial<browser.tabs.Tab>, tab: browser.tabs.Tab) {
                    bg_this.onTabUpdated(tab_id, update_info, tab);
                }
            );
        }

        public getTabData(index: number): TabData {
            if (!this.tabData[index]) {
                this.tabData[index] = new TabData(index);
            }
            return this.tabData[index];
        }

        public onMessage(message: Page_Data_Out, sender: browser.runtime.MessageSender) {
            if (sender.tab && sender.tab.id) {
                this.getTabData(sender.tab.id).onMessage(message, sender);
                // browser.pageAction.show(sender.tab.id);
            }
        }

        public onTabUpdated(tab_id: number, _update_info: Partial<browser.tabs.Tab>, _tab: browser.tabs.Tab) {
            if (this.tabData[tab_id]) {
                this.tabData[tab_id].onTabUpdated(tab_id, _update_info, _tab);
            }
        }

    }

    export type BackgroundWindow = { mjs_background: Background };


}

var mjs_background: MJS.Background = new MJS.Background();
