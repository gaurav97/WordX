var text = new String;
var searchQuery = new String;

var searchType = new String;
var prevText = new String;

(function () {
    //"use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            setInterval(getDataFromSelection, 500);
        });
    };


    // Reads data from current document sele    ction and displays a notification
    function getDataFromSelection() {
        Office.context.document.getFileAsync(Office.FileType.Text, { sliceSize: 65536 /*64 KB*/ },
        function (result) {
            if (result.status == "succeeded") {
                // If the getFileAsync call succeeded, then
                // result.value will return a valid File Object.
                var myFile = result.value;
                var sliceCount = myFile.sliceCount;
                var slicesReceived = 0, gotAllSlices = true, docdataSlices = [];

                // Get the file slices.
                getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived);

            }
            else {
                //  app.showNotification("Error:", result.error.message);
            }
        });
    }

    var a = -1, b = -1, c = -1;
    function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived) {
        file.getSliceAsync(nextSlice, function (sliceResult) {
            if (sliceResult.status == "succeeded") {
                if (!gotAllSlices) { // Failed to get all slices, no need to continue.
                    return;
                }

                // Got one slice, store it in a temporary array.
                // (Or you can do something else, such as
                // send it to a third-party server.)
                docdataSlices[sliceResult.value.index] = sliceResult.value.data;
                //$('#references').html(sliceResult.value.data);
                if (++slicesReceived == sliceCount) {
                    // All slices have been received.
                    file.closeAsync();
                    onGotAllSlices(docdataSlices);
                }
                else {
                    getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
                }
            }
            else {
                gotAllSlices = false;
                file.closeAsync();
                //  app.showNotification("getSliceAsync Error:", sliceResult.error.message);
            }
        });
    }


    function getText(a) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
                { valueFormat: "unformatted", filterType: "all" },
                function (asyncResult) {
                    var error = asyncResult.error;
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        //  write(error.name + ": " + error.message);
                    }
                    else {
                        dataValue = asyncResult.value;
                        toreplace = "::" + text + "::";
                        if (dataValue == toreplace) {
                            if (a == 1) {
                                var html = $('#wiki').html();
                                html = html.replace(/<style([\s\S]*?)<\/style>/gi, '');
                                html = html.replace(/<script([\s\S]*?)<\/script>/gi, '');
                                html = html.replace(/<\/div>/ig, '\n');
                                html = html.replace(/<\/li>/ig, '\n');
                                html = html.replace(/<li>/ig, '  *  ');
                                html = html.replace(/<\/ul>/ig, '\n');
                                html = html.replace(/<\/p>/ig, '\n');
                                html = html.replace(/<br\s*[\/]?>/gi, "\n");
                                html = html.replace(/&nbsp;/ig, '');

                                html = html.replace(/<[^>]+>/ig, '');



                                replaceWith = html;

                            }
                            else if (a == 2) {
                                var html = $('#quote').html();
                                html = html.replace(/<style([\s\S]*?)<\/style>/gi, '');
                                html = html.replace(/<script([\s\S]*?)<\/script>/gi, '');
                                html = html.replace(/<\/div>/ig, '\n');
                                html = html.replace(/<\/li>/ig, '\n');
                                html = html.replace(/<li>/ig, '  *  ');
                                html = html.replace(/<\/ul>/ig, '\n');
                                html = html.replace(/<\/p>/ig, '\n');
                                html = html.replace(/<br\s*[\/]?>/gi, "\n");
                                html = html.replace(/&nbsp;/ig, '');

                                html = html.replace(/<[^>]+>/ig, '');
                                html = "\"" + $('#quote').html() + "\"";
                                replaceWith = html;

                            }
                            Office.context.document.setSelectedDataAsync(replaceWith,
                            function (asyncResult) {
                                var error = asyncResult.error;
                                if (asyncResult.status === "failed") {
                                    //  write(error.name + ": " + error.message);
                                }
                            });
                        }
                    }
                });
    }
    function onGotAllSlices(docdataSlices) {
        var docdata = [];
        for (var i = 0; i < docdataSlices.length; i++) {
            docdata = docdata.concat(docdataSlices[i]);

        }
        prevText = text;

        a = docdata[0].indexOf("::");
        if (a > 0) {
            b = docdata[0].indexOf("::", a + 2);
            if (b > 0) {
                text = docdata[0].substring(a + 2, b);
            }
        }
        c = text.indexOf(":");
        searchType = text.substring(0, c).toLowerCase();
         searchQuery= text.substring(c + 1, text.length);


        /*var fileContent = new String();
        for (var j = 0; j < docdata.length; j++) {
            fileContent += String.fromCharCode(docdata[j]);
        }*/
        //app.showNotification(" Slices: " + fileContent);
        // Now all the file content is stored in 'fileContent' variable,
        // you can do something with it, such as print, fax...
        if (searchType == 'wiki')
            getText(1);
        else if (searchType == 'quote')
            getText(2)

        if (prevText != text) {
            if (a > 0 && b > 0) {

                if (searchType == 'img') {
                    imageSearch = new google.search.ImageSearch();

                    // Set searchComplete as the callback function when a search is
                    // complete.  The imageSearch object will have results in it.
                    imageSearch.setSearchCompleteCallback(this, searchComplete, null);

                    // Find me a beautiful car.
                    //$('#references').html(text  + searchQuery + "boo" + searchType);
                    imageSearch.execute(searchQuery); //Enter whatever you want to search for here

                    // Include the required Google branding
                    // google.search.Search.getBranding('branding');
                    $('#contentTopic').html("Image recommendations")
                    $("#wikiTopic").html("");
                    $("#quoteTopic").html("");
                    $("#wiki").html("");
                    $("#quote").html("");
                }
                else if (searchType == 'wiki' && searchQuery != '') {
                    $('#wikiTopic').html("Content recommendations");
                    $("#quoteTopic").html("");
                    $("#contentTopic").html("");
                    $("#quote").html("");
                    $("#content").html("");
                    $.getJSON("http://en.wikipedia.org/w/api.php?callback=?",
        {
            srsearch: searchQuery,
            action: "query",
            srlimit: 1,

            list: "search",
            format: "json"
        },
        function (data) {
            p = data.query.search.title;
            $.each(data.query.search, function (i, item) {
                searchQuery = encodeURIComponent(item.title);

                $('#wiki').wikiblurb({
                    wikiURL: "http://en.wikipedia.org/",
                    apiPath: 'w',
                    section: 0,
                    page: searchQuery,
                    removeLinks: true,
                    type: 'text',
                    customSelector: '',
                    callback: function () {
                        console.log('Data loaded...');
                    }
                }

              );



            });
        });




                }

                else if (searchType == 'quote') {
                    $("#content").html("");
                    $("#wiki").html("");
                    $("#contentTopic").html("");
                    $("#wikiTopic").html("");
                    $('#quoteTopic').html("Quotes");
                    quotes = {};

                    /*
                    function quoteReady(newQuote, quoteDiv, index) {
                      quoteDiv.html($('<p style="display:none;">' + newQuote + '</p>'));
                      quoteDiv.find("p:hidden").fadeIn(400);
                    }
                    */

                    function quoteReady(newQuote) {
                        if (newQuote.quote) {
                            if (newQuote.quote.length > 0) {

                                $('#quote').html(newQuote.quote);


                            }
                        }
                    }



                    WikiquoteApi.openSearch(searchQuery,
                          function (results) {

                              // Get quote
                              WikiquoteApi.getRandomQuote(results[0],
                                function (newQuote) { quoteReady(newQuote); },
                                function (msg) {

                                }
                                );
                          },
                          function (msg) {

                          }
                          );

                    getText(2);

                }
                else {
                    $("#content").html("");
                    $("#wiki").html("");
                    $('#quote').html("");
                    $("#contentTopic").html("");
                    $("#wikiTopic").html("");
                    $('#quoteTopic').html("");

                }
            }

        }
    }
    //imagesearch("asd");




})();