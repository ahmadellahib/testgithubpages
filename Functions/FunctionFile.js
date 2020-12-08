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
   // var uri = "https://qmdevstorageaccount.blob.core.windows.net/sc-container/ahmad.ellahib@strategy-compass.com-new";
//     uri = "https://qmdevstorageaccount.blob.core.windows.net/sc-container/testjs.js";
//    // uri = "https://api.qmdev2020.com/api/values";


//     // method 1
//     // $.ajax({
//     //     url: uri,
//     //     type:'GET',
//     //     dataType: "jsonp",
//     //     crossDomain:true,
//     //     beforeSend: function (request) {
//     //         request.setRequestHeader("Access-Control-Allow-Origin", "*");
//     //     },
//     //     success: function (data) {
//     //         console.log("Ahmad",data); 
//     //         signature = "signature is found";
//     //     },
//     //     error: function (xhr, textStatus, errorMessage) {
//     //         console.log("Ahmad",errorMessage); 
//     //         signature = "ERROR signature is not found";
//     //     }                
//     // });

//     // method 2
//     var script = document.createElement("script");
//     script.setAttribute("src", uri);
//     document.getElementsByTagName('head')[0].appendChild(script);

console.log("Ahmad1",item);
console.log("Ahmad2",item.new);

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