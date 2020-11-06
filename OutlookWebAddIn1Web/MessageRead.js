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
        body.getAsync(Office.CoercionType.Text, function (asyncResult) {
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.log(asyncResult);
                console.log("Failed");
            }
            else {
                $('#item-internetMessageId').html(asyncResult.value.trim());
            }
        });       
        
        // Write message property values to the task pane
        $('#item-id').text(item.itemId);
        $('#item-subject').text(item.subject);
        $('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");
    }
})();
