/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */


/* Adding the script tag to the head */

var head = document.getElementsByTagName('head')[0];
var script = document.createElement('script');
script.type = 'text/javascript';
script.src = "https://code.jquery.com/jquery-3.3.1.min.js";

// Then bind the event to the callback function.
// There are several events for cross browser compatibility.
script.onreadystatechange = handler;
script.onload = handler;

// Fire the loading
head.appendChild(script);

function handler() {
    console.log('jquery added :)');
}

var mailboxItem;

Office.initialize = function (reason) {
    mailboxItem = Office.context.mailbox.item;
}

// Entry point for Contoso Message Body Checker add-in before send is allowed.
// <param name="event">MessageSend event is automatically passed by BlockOnSend code to the function specified in the manifest.</param>
function validateBody(event) {
    mailboxItem.body.getAsync(Office.CoercionType.Text, { asyncContext: event }, checkBodyOnlyOnSendCallBack);
   
}

// Invoke by Contoso Subject and CC Checker add-in before send is allowed.
// <param name="event">MessageSend event is automatically passed by BlockOnSend code to the function specified in the manifest.</param>
function validateSubjectAndCC(event) {
      shouldChangeSubjectOnSend(event);
   
    //appendText(event);
}

function appendText(event) {
    console.log('Inside append ');
    debugger;
    Office.context.mailbox.item.body.setAsync(
        '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
        { coercionType: Office.CoercionType.Html }
    );
    Office.context.mailbox.send();
    //var newHtml = event.value.replace("</body>", "<br/ >apend text here.</body>")
    ////    Office.context.mailbox.item.body.setAsync(newHtml, { coercionType: Office.CoercionType.Html });
    //mailboxItem.body.setAsync('TESTSTSTSTSSTSST', { coercionType: Office.CoercionType.Text }, { asyncContext: event });
    //asyncResult.asyncContext.completed({ allowEvent: true });
    
    //Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function (result) {
    //    var newHtml = result.value.replace("</body>", "<br/ >apend text here.</body>")
    //    Office.context.mailbox.item.body.setAsync(newHtml, { coercionType: Office.CoercionType.Html });
    //});
}

//function addBodyContents(event) {
//    var newHtml = event.value.replace("</body>", "<br/ >apend text here.</body>")
//    //    Office.context.mailbox.item.body.setAsync(newHtml, { coercionType: Office.CoercionType.Html });
//    mailboxItem.body.setAsync('TESTSTSTSTSSTSST', { coercionType: Office.CoercionType.Text }, { asyncContext: event });
//    asyncResult.asyncContext.completed({ allowEvent: true });
//}


// Check if the subject should be changed. If it is already changed allow send. Otherwise change it.
// <param name="event">MessageSend event passed from the calling function.</param>
function shouldChangeSubjectOnSend(event) {
    //Office.context.mailbox.item.body.getAsync(
    //          Office.CoercionType.Text,
    //          { asyncContext: event },
    //          function callback(result) {
    //              // Do something with the result
    //              result = 'Change in body ' + result.asyncContext;
    //              console.log('Change in body ' + result.asyncContext);
    //          });


    mailboxItem.subject.getAsync(
        { asyncContext: event },
        function (asyncResult) {
            addCCOnSend(asyncResult.asyncContext);
            //console.log(asyncResult.value);
            // Match string.
            var checkSubject = (new RegExp(/\[Checked\]/)).test(asyncResult.value)
            // Add [Checked]: to subject line.

      
            console.log('API Call start');
            var request = new XMLHttpRequest();
            request.open("get", "https://addinwebapinew.azurewebsites.net/api/OWA/" + asyncResult.value + "/GetUsers", false);
            request.send();
            console.log("Response Text " + request.responseText);
            var myResult = request.responseText;
            console.log('API Call end ');
            //subject = '[Checked]: ' + asyncResult.value + myResult;

            var checkVal = localStorage.getItem("checkValue");
            console.log('CheckVal' + checkVal);
            if (checkVal == 'true') {
                console.log('Inside True ');
                subject = '[Encrypted]: ' + asyncResult.value + myResult;
                localStorage.clear();
            }
            else {
                console.log('Inside false');
                subject = asyncResult.value;
                localStorage.clear();
            }
            // Check if a string is blank, null or undefined.
            // If yes, block send and display information bar to notify sender to add a subject.
            if (asyncResult.value === null || (/^\s*$/).test(asyncResult.value)) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Please enter a subject for this email.' });
                asyncResult.asyncContext.completed({ allowEvent: false });
               }
            else {
                // If can't find a [Checked]: string match in subject, call subjectOnSendChange function.
                if (!checkSubject) {
                    subjectOnSendChange(subject, asyncResult.asyncContext);
                   
                    //console.log(checkSubject);
                 }
                else {
                    // Allow send.
                    asyncResult.asyncContext.completed({ allowEvent: true });
                 }
            }

        }
    )
}

// Add a CC to the email.  In this example, CC contoso@contoso.onmicrosoft.com
// <param name="event">MessageSend event passed from calling function</param>
function addCCOnSend(event) {
    mailboxItem.cc.setAsync(['Contoso@contoso.onmicrosoft.com'], { asyncContext: event });
}

// Check if the subject should be changed. If it is already changed allow send, otherwise change it.
// <param name="subject">Subject to set.</param>
// <param name="event">MessageSend event passed from the calling function.</param>
function subjectOnSendChange(subject, event) {
    mailboxItem.body.setAsync(
        subject,
        { asyncContext: event },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Unable to set the body.' });

                // Block send.
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                // Allow send.

                asyncResult.asyncContext.completed({ allowEvent: true });
            }

        });

   
}

function bodyOnSendChange(subject, event) {
    mailboxItem.body.setAsync(
        subject,
        { asyncContext: event },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Unable to set the body.' });

                // Block send.
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                // Allow send.

                asyncResult.asyncContext.completed({ allowEvent: true });
            }

        });


}


// Check if the body contains a specific set of blocked words. If it contains the blocked words, block email from being sent. Otherwise allows sending.
// <param name="asyncResult">MessageSend event passed from the calling function.</param>
function checkBodyOnlyOnSendCallBack(asyncResult) {
    var listOfBlockedWords = new Array("blockedword", "blockedword1", "blockedword2");
    var wordExpression = listOfBlockedWords.join('|');

    // \b to perform a "whole words only" search using a regular expression in the form of \bword\b.
    // i to perform case-insensitive search.
    var regexCheck = new RegExp('\\b(' + wordExpression + ')\\b', 'i');
    var checkBody = regexCheck.test(asyncResult.value);
    console.log('checkBody' + checkBody);
    if (checkBody) {
        mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Blocked words have been found in the body of this email. Please remove them.' });
        // Block send.
        console.log('AsyncResult block' + asyncResult.value);
        mailboxItem.body.setAsync('<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>', { coercionType: Office.CoercionType.Html });
        asyncResult.asyncContext.completed({ allowEvent: false });
        console.log('API Call start');
        var request = new XMLHttpRequest();
        request.open("get", "https://addinwebapinew.azurewebsites.net/api/OWA/" + asyncResult.value + "/GetUsers", false);
        request.send();
        console.log("Response Text " +  request.responseText);
        console.log('API Call end ');
    }
         // Allow send.
    console.log('AsyncResult allow ' + asyncResult.value);
    var request = new XMLHttpRequest();
    request.open("get", "https://addinwebapinew.azurewebsites.net/api/OWA/" + asyncResult.value + "/GetUsers", false);
    request.send();
    console.log("Response Text " + request.responseText);
   // mailboxItem.body.setAsync('<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>', { coercionType: Office.CoercionType.Html });
    asyncResult.asyncContext.completed({ allowEvent: true });
}
