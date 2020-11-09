'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            loadItemProps(Office.context.mailbox.item);
            setItemBody(Office.context.mailbox.item);
        });
    });

    function setItemBody(item) {

                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of setSelectedDataAsync.
                    item.body.setSelectedDataAsync(
                        '<b> Kindly note we now open 7 days a week.</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
//                 }
//                 else {
//                     // Body is of text type. 
//                     item.body.setSelectedDataAsync(
//                         ' Kindly note we now open 7 days a week.',
//                         { coercionType: Office.CoercionType.Text, 
//                             asyncContext: { var3: 1, var4: 2 } },
//                         function (asyncResult) {
//                             if (asyncResult.status == 
//                                 Office.AsyncResultStatus.Failed){
//                                 write(asyncResult.error.message);
//                             }
//                             else {
//                                 // Successfully set data in item body.
//                                 // Do whatever appropriate for your scenario,
//                                 // using the arguments var3 and var4 as applicable.
//                             }
//                          });
//                 }
//             }
//         });
}
    
   function write(message){
    document.getElementById('message').innerText += message; 
   }
    
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
})();
