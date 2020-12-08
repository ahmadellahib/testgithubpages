Office.initialize = function () {
    
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
        statusUpdate("icon16", "Setting subject done successfully!");
    });
}

function setSignature() {    
    console.log("Ahmad1",item);
    console.log("Ahmad2",item.new);

    // Set the signature for the current item.
    var signature = item.new;

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