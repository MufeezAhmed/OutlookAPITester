Office.initialize = function () {
}

// Helper function to add a status message to the info bar.
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

function defaultStatus(event) {
  statusUpdate("icon16" , "Hello World!");
}


function displayReplyForm(event) {
    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
    /* Display a reply form */
    Office.context.mailbox.item.displayReplyForm(
        {
            'htmlBody': 'hi',
            'attachments': [
                {
                    'type': Office.MailboxEnums.AttachmentType.File,
                    'name': 'squirrel.png',
                    'url': 'http://i.imgur.com/sRgTlGR.jpg'
                },
                {
                    'type': Office.MailboxEnums.AttachmentType.Item,
                    'name': 'mymail',
                    'itemId': Office.context.mailbox.item.itemId
                }
            ]
        }
    );

    event.completed();
}

function displayReplyAllForm(event) {
    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
    /* Display a reply all form */
    Office.context.mailbox.item.displayReplyAllForm("hi");
    event.completed();
}

function displayReplyFormWithInline(event) {
    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
    /* inline image - display reply form */
    Office.context.mailbox.item.displayReplyForm(
        {
            'htmlBody': '<img src = "cid:squirrel.png">',
            'attachments':
            [
                {
                    'type': Office.MailboxEnums.AttachmentType.File,
                    'name': 'squirrel.png',
                    'url': 'http://i.imgur.com/sRgTlGR.jpg',
                    'isInline': 'true'
                }
            ]
        });
    event.completed();
}

function displayReplyAllFormWithInline(event) {
    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
    /* inline image - display reply form */
    Office.context.mailbox.item.displayReplyAllForm(
        {
            'htmlBody': '<img src = "cid:squirrel.png">',
            'attachments':
            [
                {
                    'type': Office.MailboxEnums.AttachmentType.File,
                    'name': 'squirrel.png',
                    'url': 'http://i.imgur.com/sRgTlGR.jpg',
                    'isInline': 'true'
                }
            ]
        });
    event.completed();
}

function displayNewAppointmentForm(event)
{
    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
    /* Display new appointment form */
    var start = new Date();
    var end = new Date();
    end.setHours(start.getHours() + 1);

    Office.context.mailbox.displayNewAppointmentForm(
        {
            requiredAttendees: ["bob@contoso.com"],
            optionalAttendees: ["sam@contoso.com"],
            start: start,
            end: end,
            location: "Home",
            resources: ["projector@contoso.com"],
            subject: "meeting",
            body: "Hello World!"
        });

    event.completed();
}

function displayMessageForm(event)
{
    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
    /* Display message form */
    // Item ID of current message
    var messageId = "AAMkAGZiZjc1Y2RkLTczNjktNGU1YS1hYTkzLTYzZTU3OTE5OWQ3NABGAAAAAAC3Bc26XexrR4XknrAwz6j9BwBDfaKHIE1iQJlAjLUe7EC6AACETbQlAABDfaKHIE1iQJlAjLUe7EC6AACI3qQ2AAA=";
    Office.context.mailbox.displayMessageForm(messageId);
    event.completed();
}

function displayAppointmentForm(event)
{
    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
    /* Display appointment form */
    // Item ID of current appointment
    var appointmentId = "AAMkAGZiZjc1Y2RkLTczNjktNGU1YS1hYTkzLTYzZTU3OTE5OWQ3NABGAAAAAAC3Bc26XexrR4XknrAwz6j9BwBDfaKHIE1iQJlAjLUe7EC6AAAAAAENAABDfaKHIE1iQJlAjLUe7EC6AABzLImxAAA=";
    Office.context.mailbox.displayAppointmentForm(appointmentId);
    event.completed();

}

function closeTaskPane(event)
{

    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
    /* close Container */
    Office.context.ui.closeContainer()//;

    event.completed();

}

function displayWebDialog(event)
{
    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
    /* displayDialog */
    var dialogOptions = { height: 80, width: 50, displayInIframe: false, requireHTTPS: false };

    Office.context.ui.displayDialogAsync("https://trelloaddin.azurewebsites.net/trello/LoginPageIOS.html", dialogOptions, displayDialogCallback);



    function displayDialogCallback(asyncResult) {

        console.log(asyncResult.status);

        expect(asyncResult.status).toBe("succeeded");
        done();
    }

    event.completed();


}
