Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    Office.actions.associate('showTaskpane', showTaskpane);
  }
});

function showTaskpane() {
  Office.context.ui.displayDialogAsync(
    window.location.origin + '/src/taskpane.html',
    { height: 60, width: 30, displayInIframe: true },
    () => {}
  );
}
