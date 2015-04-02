/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#create-link').click(getDataFromSelection);
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                	var text = result.value;
                	var hyperlink = '<a href="http://www.bing.com/search?q=' + text + '">' + text + '</a>';
                	setDataSelection(hyperlink);
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }

    function setDataSelection(hyperText) {
    	Office.context.document.setSelectedDataAsync(hyperText, { coercionType: 'html' },
			function (result) {
				if (result.status === Office.AsyncResultStatus.Succeeded) {
					app.showNotification('Success!', 'Link created!');
				} else {
					app.showNotification('Error:', result.error.message);
				}
		});
    }
})();