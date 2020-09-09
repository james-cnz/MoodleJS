/**
 * Moodle JS Background Script
 * Links to data for each tab, and passes on messages from content scripts.
 */


namespace MJS {

    export type BackgroundWindow = { mjs_background: Background };

    export class Background {

        private tabData: { [index: number]: TabData } = { };

        public constructor() {
            const this_bg = this;
            browser.runtime.onMessage.addListener(
                function(message: object, sender: browser.runtime.MessageSender) {
                    this_bg.onMessage(message as Page_Data, sender);
                }
            );
            browser.tabs.onUpdated.addListener(
                function(tab_id: number, update_info: Partial<browser.tabs.Tab>, tab: browser.tabs.Tab) {
                    this_bg.onTabUpdated(tab_id, update_info, tab);
                }
            );
        }

        public getTabData(index: number): TabData {
            if (!this.tabData[index]) {
                this.tabData[index] = new TabData(index);
            }
            return this.tabData[index];
        }

        public onMessage(message: Page_Data, sender: browser.runtime.MessageSender) {
            if (sender.tab && sender.tab.id) {
                this.getTabData(sender.tab.id).onMessage(message, sender);
            }
        }

        public onTabUpdated(tab_id: number, update_info: Partial<browser.tabs.Tab>, tab: browser.tabs.Tab) {
            if (this.tabData[tab_id]) {
                this.tabData[tab_id].onTabUpdated(tab_id, update_info, tab);
            }
        }

    }


}

// tslint:disable-next-line: no-var-keyword
var mjs_background: MJS.Background = new MJS.Background();
