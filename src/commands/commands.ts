export function registerCommandHandlers(): void {
  // The compose entrypoint is now taskpane-first. Outlook still loads the
  // function file, but there are no execute-function commands to associate here.
}

if (typeof Office !== "undefined") {
  void Office.onReady(() => {
    registerCommandHandlers();
  });
}
