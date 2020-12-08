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
    // Set the signature for the current item.
    var signature = "This is my qm signature";
    var uri = "https://qmdevstorageaccount.blob.core.windows.net/sc-container/ahmad.ellahib@strategy-compass.com-new";
    uri = "https://qmdevstorageaccount.blob.core.windows.net/sc-container/test.json";
  
    $.ajax({
        url: uri,
        type:'GET',
        dataType: "jsonp",
        contentType: "json",
        crossDomain:true,
        beforeSend: function (request) {
            request.setRequestHeader("Access-Control-Allow-Origin", "*");
        },
        success: function (data) {
            console.log("Ahmad",data); 
        },
        error: function (xhr, textStatus, errorMessage) {
            console.log("Ahmad",errorMessage); 
        }                
    });

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