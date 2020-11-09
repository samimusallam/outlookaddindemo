'use strict';

(function () {

    var xhr;
var serviceRequest;
    
    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            initApp();
//             loadItemProps(Office.context.mailbox.item);
        });
    });
    
   function loadItemProps(item) {
//        var body = item.body;
//         $('#item-MessageBody').text("Initial Body");
//         body.getAsync(Office.CoercionType.Html, function (asyncResult) {
//             if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
//                 console.log(asyncResult);
//                 $('#item-MessageBody').text("Failed to get body");
//             }
//             else {
//                 $('#item-MessageBody').text("JSON: " + JSON.stringify(asyncResult, null, 2));
//             }
//         });    
        
        var outputString = "attachments: ";
        console.log("TESTING");
        if (item.attachments.length > 0) {
            for (var i = 0 ; i < item.attachments.length ; i++) {
                var attachment = item.attachments[i];
                outputString += "<BR>" + i + ". Name: ";
                outputString += attachment.name;
                outputString += "<BR>ID: " + attachment.id;
                outputString += "<BR>contentType: " + attachment.contentType;
                outputString += "<BR>size: " + attachment.size;
                outputString += "<BR>attachmentType: " + attachment.attachmentType;
                outputString += "<BR>isInline: " + attachment.isInline;
            }
        }
        
        $('#item-MessageBody').html(outputString);
        
        // Write message property values to the task pane
        $('#item-id').text(item.itemId);
        $('#item-subject').text(item.subject);
        $('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");
    }
    
        function initApp() {
        $("#footer").hide();

        if (Office.context.mailbox.item.attachments == undefined) {
            var testButton = document.getElementById("testButton");
            testButton.onclick = "";
            showToast("Not supported", "Attachments are not supported by your Exchange server.");
        } else if (Office.context.mailbox.item.attachments.length == 0) {
            var testButton = document.getElementById("testButton");
            testButton.onclick = "";
            showToast("No attachments", "There are no attachments on this item.");
        } else {

            // Initalize a context object for the app.
            //   Set the fields that are used on the request
            //   object to default values.
            serviceRequest = new Object();
            serviceRequest.attachmentToken = "";
            serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
            serviceRequest.attachments = new Array();
        }
    };

})();

function testAttachments() {
    Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
};

function attachmentTokenCallback(asyncResult, userContext) {
    if (asyncResult.status == "succeeded") {
        serviceRequest.attachmentToken = asyncResult.value;
        makeServiceRequest();
    }
    else {
        showToast("Error", "Could not get callback token: " + asyncResult.error.message);
    }
}

function makeServiceRequest() {
    var attachment;
    xhr = new XMLHttpRequest();

    // Update the URL to point to your service location.
    xhr.open("POST", "https://localhost:44320/api/AttachmentService", true);

    xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    xhr.onreadystatechange = requestReadyStateChange;

    // Translate the attachment details into a form easily understood by WCF.
    for (var i = 0; i < Office.context.mailbox.item.attachments.length; i++) {
        attachment = Office.context.mailbox.item.attachments[i];
        attachment = attachment._data$p$0 || attachment.$0_0;

        if (attachment !== undefined) {
            serviceRequest.attachments[i] = JSON.parse(JSON.stringify(attachment));
        }
    }

    // Send the request. The response is handled in the 
    // requestReadyStateChange function.
    xhr.send(JSON.stringify(serviceRequest));
};


// Handles the response from the JSON web service.
function requestReadyStateChange() {
    if (xhr.readyState == 4) {
        if (xhr.status == 200) {
            var response = JSON.parse(xhr.responseText);
            if (!response.isError) {
                // The response indicates that the server recognized
                // the client identity and processed the request.
                // Show the response.
                var names = "<h2>Attachments processed: " + response.attachmentsProcessed + "</h2>";

                for (var i = 0; i < response.attachmentNames.length; i++) {
                    names += response.attachmentNames[i] + "<br />";
                }
                document.getElementById("names").innerHTML = names;
            } else {
                showToast("Runtime error", response.message);
            }
        } else {
            if (xhr.status == 404) {
                showToast("Service not found", "The app server could not be found.");
            } else {
                showToast("Unknown error", "There was an unexpected error: " + xhr.status + " -- " + xhr.statusText);
            }
        }
    }
};

// Shows the service response.
function showResponse(response) {
    showToast("Service Response", "Attachments processed: " + response.attachmentsProcessed);
}

// Displays a message for 10 seconds.
function showToast(title, message) {

    var notice = document.getElementById("notice");
    var output = document.getElementById('output');

    notice.innerHTML = title;
    output.innerHTML = message;

    $("#footer").show("slow");

    window.setTimeout(() => { $("#footer").hide("slow") }, 10000);
};
})();
