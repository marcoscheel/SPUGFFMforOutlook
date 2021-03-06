﻿/// <reference path="../App.js" />

(function () {
    "use strict";
    var currentEventIndex;
    var myEvents;

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $.ajax({
                url: '../../Data/events.json',
                type: 'GET',
                dataType: 'json',
                contentType: 'application/json;charset=utf-8'
            }).done(function (data) {

                myEvents = data.Events;
                currentEventIndex = myEvents.length - 1;
                displayCurrentEvent();

                $("#next").click(function (event) {
                    currentEventIndex++;
                    if (currentEventIndex > (myEvents.length - 1)) {
                        currentEventIndex = 0;
                    }
                    if (currentEventIndex < 0) {
                        currentEventIndex = (myEvents.length - 1);
                    }

                    displayCurrentEvent();
                });

                $("#previous").click(function (event) {
                    currentEventIndex--;
                    if (currentEventIndex > (myEvents.length - 1)) {
                        currentEventIndex = 0;
                    }
                    if (currentEventIndex < 0) {
                        currentEventIndex = (myEvents.length - 1);
                    }

                    displayCurrentEvent();
                });

                $('#loading').text('Beschreibung');

            }).fail(function (status) {
                app.showNotification('Error', 'Could not communicate with the server.');
            }).done(function () {
            });
        });
    };


    function displayCurrentEvent() {
        var latestEvent = myEvents[currentEventIndex];
        $('#title').text(latestEvent.Title);
        $('#date').text(latestEvent.Date);
        $('#location').html(latestEvent.Location.replace(/\n/g, "<br />"));
        $('#location').attr('href', 'http://maps.google.com/maps?q=' + latestEvent.LocationGPS);
        $('#description').html(latestEvent.Description.replace(/\n/g, "<br />"));
        $('#eventURL').attr('href', latestEvent.EventUrl);


        var mentions = "";
        jQuery.each(latestEvent.Mentions, function (index, value) {
            mentions += "<span class=\"glyphicon glyphicon-share\"></span>  <a target=\"_blank\" href=\"http://twitter.com/" + this + "\">" + this + "</a><br/>";
        });
        $('#social').html(mentions);

        var links = '';
        jQuery.each(latestEvent.Links, function (index, value) {
            links += '<span class=\"glyphicon glyphicon-globe\"></span>  <a target=\"_blank\" href=\"' + this + '\">' + this + '</a><br/>';
        });
        $('#links').html(links);
    }

})();