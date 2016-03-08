
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
        /*Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    app.showNotification('The selected text is:', '"' + a + '"');
                    
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );*/
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
                app.showNotification("Error:", result.error.message);
            }
        });
    }

    var a = -1, b = -1, c = -1;
    text = " ";
    searchType = "def2";
    searchQuery = "Def1";
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
                app.showNotification("getSliceAsync Error:", sliceResult.error.message);
            }
        });
    }



    function onGotAllSlices(docdataSlices) {
        var docdata = [];
        for (var i = 0; i < docdataSlices.length; i++) {
            docdata = docdata.concat(docdataSlices[i]);

        }

        a = docdata[0].indexOf(";;");
        if (a > 0) {
            b = docdata[0].indexOf(";;", a + 2);
            if (b > 0) {
                text = docdata[0].substring(a + 2, b);
            }
        }
        
        if (a > 0) {
            c = text.indexOf(":");

            searchQuery = text.substring(0, c);
            searchType = text.substring(c + 1, text.length);
        }

        $('#references').html(a + " " + b + " " + text + "\n" + searchQuery + "\n" + searchType);
        /*var fileContent = new String();
        for (var j = 0; j < docdata.length; j++) {
            fileContent += String.fromCharCode(docdata[j]);
        }*/
        //app.showNotification(" Slices: " + fileContent);
        // Now all the file content is stored in 'fileContent' variable,
        // you can do something with it, such as print, fax...
    }

    //imagesearch("asd");




})();