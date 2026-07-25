export interface NativeOfficeDialogNotification {
  type: string;
  sessionId: string;
  [key: string]: unknown;
}

// Windows Office integration no longer hosts the editor inside an Office.js
// dialog. Native Word/PowerPoint add-ins commit through the local session API,
// so there is no Office.context.ui parent to notify.
export function messageOfficeParent(
  _message: NativeOfficeDialogNotification,
) {
  return false;
}
