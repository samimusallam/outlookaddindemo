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
       var body = item.body;
        $('#item-MessageBody').text("Initial Body");
        body.getAsync(Office.CoercionType.Html, function (asyncResult) {
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.log(asyncResult);
                $('#item-MessageBody').text("Failed to get body");
            }
            else {
                $('#item-MessageBody').text("JSON: " + JSON.stringify(asyncResult, null, 2));
            }
        });       
        
        // Write message property values to the task pane
        $('#item-id').text(item.itemId);
        $('#item-subject').text(item.subject);
        $('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");
    }
})();
