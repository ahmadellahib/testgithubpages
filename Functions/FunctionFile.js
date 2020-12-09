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
    var signature = "no signature found";
    const queryString = window.location.search;
    const urlParams = new URLSearchParams(queryString);
    var uri = "https://api.qmdev2020.com/api/values/" + urlParams.get('tenantid');

    $.ajax({
        url: uri,
        type:'GET',
        dataType: "json",
        success: function(data) {
            console.log("log response on success");
            console.log(data);
            signature = data.Name;
        },
        error: function (xhr, textStatus, errorMessage) {
                console.log(errorMessage); 
                signature = "ERROR signature is not found";
        },
        complete: function(data) {
            console.log("ajax response is completed"); 
            Office.context.mailbox.item.body.setSignatureAsync(signature, { coercionType: "html" }, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("setSignatureAsync succeeded");
                    statusUpdate("icon16", "Setting signature done successfully!");
                } else {
                    console.error(asyncResult.error);
                    statusUpdate("icon16", "Setting signature Failed!");
                }
            });
        }
    });
}