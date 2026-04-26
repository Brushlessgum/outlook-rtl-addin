/*
 * RTL by Default — Outlook add-in (v1.2)
 *
 * Two entry points:
 *   1. onNewMessageComposeHandler — fires automatically when a new compose
 *      window opens. Wraps the body in an RTL div.
 *   2. toggleRtlHandler — fires on click of the "Toggle RTL" ribbon button.
 *      Flips the body between RTL and LTR.
 *
 * Both handlers use a watchdog so event.completed() is GUARANTEED to fire
 * within 3 seconds, even if the body read/write hangs. This prevents the
 * "RTL by Default is working on your request..." loading bar from getting
 * stuck.
 */

// ----- helpers -----------------------------------------------------------

const RTL_OPEN = '<div dir="rtl" style="direction: rtl; text-align: right; unicode-bidi: embed;">';
const LTR_OPEN = '<div dir="ltr" style="direction: ltr; text-align: left;">';
const CLOSE    = '</div>';
const WATCHDOG_MS = 3000;

function bodyIsRtl(html) {
    return /dir\s*=\s*["']rtl["']/i.test(html);
}

// Strip our previously-added outer wrapper if present (lazy match to avoid
// catastrophic regex backtracking on large signatures).
function unwrap(html) {
    const m = html.match(/^\s*<div\s+dir=["'](?:rtl|ltr)["'][^>]*?>([\s\S]*?)<\/div>\s*$/i);
    return m ? m[1] : html;
}

function rewrap(html, isRtl) {
    return (isRtl ? RTL_OPEN : LTR_OPEN) + unwrap(html) + CLOSE;
}

// Wraps the body-set logic in a watchdog. If anything hangs or throws,
// event.completed() is still called within WATCHDOG_MS.
function applyDirection(event, isRtl) {
    let done = false;

    const finish = () => {
        if (done) return;
        done = true;
        try {
            event.completed();
        } catch (e) {
            // ignore — Outlook will time us out anyway
        }
    };

    // Safety net: guaranteed completion no matter what
    const watchdog = setTimeout(() => {
        console.warn("RTL add-in: watchdog fired — forcing completion");
        finish();
    }, WATCHDOG_MS);

    try {
        const item = Office.context.mailbox.item;
        item.body.getAsync(Office.CoercionType.Html, function (getResult) {
            try {
                if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
                    clearTimeout(watchdog);
                    finish();
                    return;
                }
                const existing = getResult.value || "<p><br></p>";
                const newBody = rewrap(existing, isRtl);

                item.body.setAsync(
                    newBody,
                    { coercionType: Office.CoercionType.Html },
                    function () {
                        clearTimeout(watchdog);
                        finish();
                    }
                );
            } catch (innerErr) {
                console.error("RTL add-in: inner error", innerErr);
                clearTimeout(watchdog);
                finish();
            }
        });
    } catch (outerErr) {
        console.error("RTL add-in: outer error", outerErr);
        clearTimeout(watchdog);
        finish();
    }
}

// ----- entry point 1: auto on new compose --------------------------------

function onNewMessageComposeHandler(event) {
    applyDirection(event, /* isRtl */ true);
}

// ----- entry point 2: manual toggle button -------------------------------

function toggleRtlHandler(event) {
    let done = false;
    const finish = () => {
        if (done) return;
        done = true;
        try { event.completed(); } catch (e) { /* ignore */ }
    };

    const watchdog = setTimeout(() => {
        console.warn("RTL toggle: watchdog fired");
        finish();
    }, WATCHDOG_MS);

    try {
        const item = Office.context.mailbox.item;
        item.body.getAsync(Office.CoercionType.Html, function (getResult) {
            try {
                if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
                    clearTimeout(watchdog);
                    finish();
                    return;
                }
                const existing = getResult.value || "<p><br></p>";
                const flipTo = !bodyIsRtl(existing);
                const newBody = rewrap(existing, flipTo);

                item.body.setAsync(
                    newBody,
                    { coercionType: Office.CoercionType.Html },
                    function () {
                        clearTimeout(watchdog);
                        finish();
                    }
                );
            } catch (innerErr) {
                console.error("RTL toggle: inner error", innerErr);
                clearTimeout(watchdog);
                finish();
            }
        });
    } catch (outerErr) {
        console.error("RTL toggle: outer error", outerErr);
        clearTimeout(watchdog);
        finish();
    }
}

// ----- registration ------------------------------------------------------

// MUST be at top level. Office discovers handlers by name during runtime init.
if (typeof Office !== "undefined") {
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
    Office.actions.associate("toggleRtlHandler", toggleRtlHandler);
}
