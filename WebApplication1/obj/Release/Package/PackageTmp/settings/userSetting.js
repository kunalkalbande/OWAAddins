(function () {
    "use strict";

    var messageBanner;

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            debugger;
            //var element = document.querySelector('.ms-MessageBanner');
            //messageBanner = new fabric.MessageBanner(element);
            //messageBanner.hideBanner();
            // loadProps();
            //$('#btnSave').click($('#subject').text("TESTSTSTSTSTSTSSTSTS"));
            //$("#btnSave").click(function () {
            //    $('#subject').text("TESTSTSTSTSTSTSSTSTS")
            //}); 

            $('#btnSave').click(function () {
                //showNotification("HI", Office.context.mailbox.item.subject);
                // window.location.href ='https://outlook.office.com/';
                // Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (result) { console.log(result.value) })
                //Office.context.mailbox.item.body.getAsync(
                //    "text",
                //    { asyncContext: "This is passed to the callback" },
                //    function callback(result) {
                //        console.log('HH' + result.value);
                //        const newLocal = $('#lblPassword').text(result.value);
                //    });
                console.log('Inside button click');
                var usernameVal = $('#lblUsername').val();
                console.log(usernameVal);
                //Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (result) { $.get('http://addinwebapinew.azurewebsites.net/api/owa/'+  +'/GetUsers'); console.log(result.value) })
                console.log('API Call start');
                var request = new XMLHttpRequest();
                request.open("get", "https://addinwebapinew.azurewebsites.net/api/OWA/" + usernameVal + "/GetUsers", false);
                request.send();
                console.log("Response Text " + usernameVal);
                console.log('API Call end ');
                //window.location.href('https://outlook.office.com/');
                $(location).attr('href', 'https://outlook.office.com/')
            });





        });


    };




    //$('#btnSave').click($('#subject').text("TESTSTSTSTSTSTSSTSTS"));
    //$('#btnSave').click(Office.context.mailbox.item.from.text());

    // Take an array of AttachmentDetails objects and build a list of attachment names, separated by a line-break.
    function buildAttachmentsString(attachments) {
        if (attachments && attachments.length > 0) {
            var returnString = "";

            for (var i = 0; i < attachments.length; i++) {
                if (i > 0) {
                    returnString = returnString + "<br/>";
                }
                returnString = returnString + attachments[i].name;
            }

            return returnString;
        }

        return "None";
    }

    // Format an EmailAddressDetails object as
    // GivenName Surname <emailaddress>
    function buildEmailAddressString(address) {
        return address.displayName + " &lt;" + address.emailAddress + "&gt;";
    }

    // Take an array of EmailAddressDetails objects and
    // build a list of formatted strings, separated by a line-break
    function buildEmailAddressesString(addresses) {
        if (addresses && addresses.length > 0) {
            var returnString = "";

            for (var i = 0; i < addresses.length; i++) {
                if (i > 0) {
                    returnString = returnString + "<br/>";
                }
                returnString = returnString + buildEmailAddressString(addresses[i]);
            }

            return returnString;
        }

        return "None";
    }

    // Load properties from the Item base object, then load the
    // message-specific properties.
    function loadProps() {
        var item = Office.context.mailbox.item;

        $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
        $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
        $('#itemClass').text(item.itemClass);
        $('#itemId').text(item.itemId);
        $('#itemType').text(item.itemType);

        $('#message-props').show();

        $('#attachments').html(buildAttachmentsString(item.attachments));
        $('#cc').html(buildEmailAddressesString(item.cc));
        $('#conversationId').text(item.conversationId);
        $('#from').html(buildEmailAddressString(item.from));
        $('#internetMessageId').text(item.internetMessageId);
        $('#normalizedSubject').text(item.normalizedSubject);
        $('#sender').html(buildEmailAddressString(item.sender));
        $('#subject').text(item.subject);
        $('#to').html(buildEmailAddressesString(item.to));

    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();