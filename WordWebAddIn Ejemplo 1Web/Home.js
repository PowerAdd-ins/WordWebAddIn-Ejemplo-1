'use strict';

(function () {

    // The initialize function is run each time the page is loaded.
    Office.initialize = function (reason) {

        $(document).ready(function () {

            // Use this to check whether the API is supported in the Word client.
            if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                // Do something that is only available via the new APIs
                $('#button1').click(show_Welcome);
                //$('#button2').click(clear);
                $('#supportedVersion').html('This code is using Word 2016 or greater.');
            }

            else {
                // Just letting you know that this code will not work with your version of Word.
                $('#supportedVersion').html('This code requires Word 2016 or greater.');
            }
        });
    };

    function show_Welcome() {
        Word.run(function (context) {

            var thisDocument = context.document;

            var range = thisDocument.getSelection();

            //thisDocument.body.clear();
            range.insertText('"Bienvenidos a nuestro curso de Office Add-ins!!!"\n', Word.InsertLocation.start);
            range.font.size = 30;

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }


    function clear() {
        Word.run(function (context) {

            var thisDocument = context.document;

            var range = thisDocument.getSelection();

            thisDocument.body.clear();

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

})();
