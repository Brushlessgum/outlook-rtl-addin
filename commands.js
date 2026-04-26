/*
 * RTL by Default — event-based Outlook add-in handler.
 *
 * This script runs automatically every time you open a new compose window
 * (new email, reply, reply-all, or forward). It wraps the message body in
 * a <div dir="rtl"> with right alignment so the cursor and text default to
 * right-to-left — matching Windows Outlook's "default text direction" option.
 */

// The function name MUST match the FunctionName in manifest.xml
//   <LaunchEvent Type="OnNewMessageCompose" FunctionName="onNewMessageComposeHandler"/>
function onNewMessageComposeHandler(event) {
    const item = Office.context.mailbox.item;

    // Read existing body (could already contain a signature)
    item.body.getAsync(Office.CoercionType.Html, function (getResult) {
        if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("RTL add-in: failed to read body", getResult.error);
            event.completed();
            return;
        }

        const existingBody = getResult.value || "<p><br></p>";

        // Wrap whatever's there (signature included) in an RTL container.
        // The inline style is necessary because some Outlook versions strip
        // the dir attribute when the user starts typing if no style is set.
        const rtlBody =
            '<div dir="rtl" style="direction: rtl; text-align: right; unicode-bidi: embed;">' +
                existingBody +
            '</div>';

        item.body.setAsync(
            rtlBody,
            { coercionType: Office.CoercionType.Html },
            function (setResult) {
                if (setResult.status !== Office.AsyncResultStatus.Succeeded) {
                    console.error("RTL add-in: failed to set body", setResult.error);
                }
                // Always call event.completed() — otherwise Outlook hangs the compose window.
                event.completed();
            }
        );
    });
}

// Register the handler with Office so the runtime can find it by name.
// This call MUST happen at script load time.
if (typeof Office !== "undefined") {
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}
