/// <reference path="../App.js" />

(function () {
    'use strict';
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


            $('#set-event').click(setSubject);
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


    function setSubject() {
        var latestEvent = myEvents[currentEventIndex];
        var item = Office.cast.item.toItemCompose(Office.context.mailbox.item);
        item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed) {
                write(result.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Set data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of setSelectedDataAsync.
                    var mentions = "";
                    jQuery.each(latestEvent.Mentions, function (index, value) {
                        mentions += "<a target=\"_blank\" href=\"http://twitter.com/" + this + "\">" + this + "</a> ";
                    });


                    item.body.setSelectedDataAsync(
                        '<b>' + latestEvent.Title + '</b><br/>' + latestEvent.Description + '<br/><a href="' + latestEvent.EventUrl + ' " target="_blank">Zur Eventseite</a><br/><br/>Erwähnt uns auf twitter:<br/>' + mentions,
                        {
                            coercionType: Office.CoercionType.Html,
                            asyncContext: { var3: 1, var4: 2 }
                        },
                        function (asyncResult) {
                            if (asyncResult.status ==
                                Office.AsyncResultStatus.Failed) {
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    var mentions = "";
                    jQuery.each(latestEvent.Mentions, function (index, value) {
                        mentions += this + " ";
                    });
                    // Body is of text type. 
                    item.body.setSelectedDataAsync(
                        '' + latestEvent.Title + '\n\n' + latestEvent.Description + '\n\nZur Eventseite\n' + latestEvent.EventUrl + '\n\nErwähnt uns auf twitter:\n' + mentions,
                        {
                            coercionType: Office.CoercionType.Text,
                            asyncContext: { var3: 1, var4: 2 }
                        },
                        function (asyncResult) {
                            if (asyncResult.status ==
                                Office.AsyncResultStatus.Failed) {
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
            }
        });
        
    }


})();