Office.initialize = function () {

    // Create three notifications, each with a different key
    //Office.context.mailbox.item.notificationMessages.addAsync("progress", {
    //    type: "progressIndicator",
    //    message: "An add-in is processing this message."
    //});
    //Office.context.mailbox.item.notificationMessages.addAsync("information", {
    //    type: "informationalMessage",
    //    message: "The add-in processed this message.",
    //    icon: "iconid",
    //    persistent: false
    //});
    //Office.context.mailbox.item.notificationMessages.addAsync("error", {
    //    type: "errorMessage",
    //    message: "The add-in failed to process this message."
    //});

}

// Helper function to add a status message to the info bar.
function statusUpdate(icon, text) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
        type: "informationalMessage",
        icon: icon,
        message: text,
        persistent: false
    });
}

function defaultStatus(event) {
    statusUpdate("icon16", "Hello World!");
}

function getSubject() {
    statusUpdate("icon16", `Emil subject ${Office.context.mailbox.item.subject}`);
}

function setSubject() {
    let subject = "This is my default subject";

    Office.context.mailbox.item.subject.setAsync(subject, (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.error(`Action failed with message ${result.error.message}`);
            return;
        }
        console.log(`Successfully set subject to ${subject}`);

        statusUpdate("icon16", "Setting subject done successfully!");

        //Office.context.mailbox.item.notificationMessages.getAllAsync(function (asyncResult) {
        //    if (asyncResult.status != "failed") {
        //        Office.context.mailbox.item.notificationMessages.replaceAsync("notifications", {
        //            type: "informationalMessage",
        //            message: "Found " + asyncResult.value.length + " notifications.",
        //            icon: "iconid",
        //            persistent: false
        //        });
        //    }
        //});


    });
}

function setSignature() {
    // Set the signature for the current item.
    var signature = "This is my qm signature";
    statusUpdate("icon16", "Ahmad Setting signature to ...");
    console.log(`Setting signature to "${signature}".`);
    Office.context.mailbox.item.body.setSignatureAsync(signature, { coercionType: "html" }, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("setSignatureAsync succeeded");
            statusUpdate("icon16", "Ahmad Setting signature done successfully!");
        } else {
            console.error(asyncResult.error);
            statusUpdate("icon16", "Ahmad Setting signature Failed!");
        }
    });
}