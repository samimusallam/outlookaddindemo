'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            loadItemProps(Office.context.mailbox.item);
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

        if (item.attachments.length > 0) {
            for (i = 0 ; i < item.attachments.length ; i++) {
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
})();
