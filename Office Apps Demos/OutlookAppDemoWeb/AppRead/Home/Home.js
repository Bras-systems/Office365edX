/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            displayItemDetails();
        });
    };

    // Displays the "Subject" and "From" fields, based on the current mail item
    function displayItemDetails() {
        var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
        var meetingSuggestions = item.getEntitiesByType(Office.MailboxEnums.EntityType.MeetingSuggestion);
        var meetingSuggestion = meetingSuggestions[0];

        $('#start-time').text(meetingSuggestion.start);
    }
})();