(async () => {
    const src = chrome.runtime.getURL("page.js");
    const contentMain = await import(src);
    // await contentMain.page_init();
  })();
