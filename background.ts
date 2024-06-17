/**
 * Moodle JS Background Script
 * Links to data for each tab, and passes on messages from content scripts.
 */

import "./browser_polyfill_mv3.js";
import {DeepPartial} from "./shared.js";
import {TabData} from "./macros.js";
import {Page_Data} from "./page.js";


    export type BackgroundWindow = { mjs_background: Background };

    export class Background {

        private tabData: { [index: number]: TabData } = { };

        public constructor() {
            browser.runtime.onMessage.addListener(
                (message: object, sender: browser.runtime.MessageSender) => {
                    this.onMessage(message as Page_Data, sender);
                }
            );
            browser.tabs.onUpdated.addListener(
                (tab_id: number, update_info: Partial<browser.tabs.Tab>, tab: browser.tabs.Tab) => {
                    this.onTabUpdated(tab_id, update_info, tab);
                }
            );
            browser.runtime.onConnect.addListener(
                (port: browser.runtime.Port) => {
                    this.connected(port);
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
            this.getTabData(tab_id).onTabUpdated(tab_id, update_info, tab);
        }

        public connected(port: browser.runtime.Port) {
            const index: number = parseInt(port.name.match(/^popup (\d+)$/)![1]);
            this.getTabData(index).popupPort = port;
            port.onMessage.addListener((message: DeepPartial<TabData>) => {
                this.getTabData(index).onPopupMessage(message);
            });
            this.getTabData(index).popup_init2();
        }

    }



// eslint-disable-next-line no-var, @typescript-eslint/no-unused-vars
var mjs_background: Background = new Background();
