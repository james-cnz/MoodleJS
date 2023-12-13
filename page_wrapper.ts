(async () => {
    const src = chrome.runtime.getURL("page.js");
    await import(src);
  })();
