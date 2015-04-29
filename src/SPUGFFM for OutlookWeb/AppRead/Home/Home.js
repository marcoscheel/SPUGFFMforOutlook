/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $.ajax({
                url: '../../Data/events.json',
                type: 'GET',
                dataType: 'json',
                contentType: 'application/json;charset=utf-8'
            }).done(function (data) {

                $('#title').text(data.Events[1].Title);
                $('#loading').hide();

            }).fail(function (status) {
                app.showNotification('Error', 'Could not communicate with the server.');
            }).done(function () {
            });
        });
    };

})();