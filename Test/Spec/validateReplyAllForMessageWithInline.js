Office.initialize = function (reason) { };

jasmine.DEFAULT_TIMEOUT_INTERVAL = 15000;


function sendTestReport() {

    var testResults = $("#testResults").html();

    $("#testResults").remove();


    $(".jasmine_html-reporter").after(testResults);


    var completeHtml = "<html>" + $("html").html() + "</html>";



    var textcompleteHtml = completeHtml;
    textcompleteHtml = textcompleteHtml.replace(/&/g, '&amp;');
    textcompleteHtml = textcompleteHtml.replace(/</g, '&lt;');
    textcompleteHtml = textcompleteHtml.replace(/>/g, '&gt;');



    var options = {
        isRest: true,
        asyncContext: { message: 'Hello World!' }
    };

    Office.context.mailbox.getCallbackTokenAsync(options, cb);


    function cb(asyncResult) {
        var token = asyncResult.value;

        var sendMessageUrl = Office.context.mailbox.restUrl +
            '/v2.0/me/sendmail';

        var emailData = {
            "Message": {
                "Subject": "Compose API Test Result for " + Office.context.mailbox.diagnostics.hostName + ":" + Office.context.mailbox.diagnostics.hostVersion,
                "Body": {
                    "ContentType": "Html",
                    "Content": completeHtml
                },
                "ToRecipients": [
                    {
                        "EmailAddress": {
                            "Address": Office.context.mailbox.userProfile.emailAddress
                        }
                    }
                ],
                "Attachments": [
                ]
            },
            "SaveToSentItems": "false"
        };

        $.ajax({
            url: sendMessageUrl,
            contentType: 'application/json',
            data: JSON.stringify(emailData),
            type: 'post',
            headers: { 'Authorization': 'Bearer ' + token }
        }).done(function (item) {
           
        }).fail(function (error) {
            $(".jasmine_html-reporter").after("<p>" + error + "</p>");
            console.log(error);
        });




    }


}

describe("",
    function () {
        beforeAll(function (done) { setTimeout(function () { done(); }, 2000) });
        afterAll(function () {


            setTimeout(function () {
                sendTestReport();
            }, 5000)


        })
     
         describe("Office.context.mailbox.item", function () {

           
            

              



                it("Get body type Async",
                    function (done) {


                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get body type */
                        Office.context.mailbox.item.body.getTypeAsync(
                            function (asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                    document.getElementById("getBodyTypeAsync").innerHTML = "Action failed with error: " + asyncResult.error.message;
                                } else {
                                    console.log(asyncResult.value);
                                    document.getElementById("getBodyTypeAsync").innerHTML = asyncResult.value
                                }
                                expect(asyncResult.status).toBe("succeeded");
                                 expect(asyncResult.value).toBe("html");
                                done();
                            }
                        );




                    });

               

               
               
                it("Get Subject Async",
                    function (done) {


                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get subject */
                        Office.context.mailbox.item.subject.getAsync(
                            function (asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                    document.getElementById("getSubjectAsync").innerHTML = "Action failed with error: " + asyncResult.error.message;
                                } else {
                                    console.log(asyncResult.value);
                                    document.getElementById("getSubjectAsync").innerHTML = asyncResult.value;
                                }
                                expect(asyncResult.status).toBe("succeeded");
                                expect(asyncResult.value).toBe("New subject!"); 
                                done();
                            }
                        );




                    });

               



               

                    it("Get To recipients ",
                        function (done) {


                            /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                            /* Get To recipients (Applicable only on message) */
                            Office.context.mailbox.item.to.getAsync(
                                function callback(asyncResult) {
                                    if (asyncResult.status == "failed") {
                                        console.log("Action failed with error: " + asyncResult.error.message);
                                        document.getElementById('getToRecipients').innerHTML = ("Action failed with error: " + asyncResult.error.message);
                                    } else {
                                        var recipients = "";
                                        asyncResult.value.forEach(function (recipient, index) {
                                            recipients = recipients + recipient.displayName + " (" + recipient.emailAddress + ");";
                                        });
                                        console.log(recipients);
                                        document.getElementById('getToRecipients').innerHTML = (recipients);
                                      
                                    }
                                    expect(asyncResult.status).toBe("succeeded");
                                      expect(recipients).toBe("Allie Bellew (allieb@contoso.com);Alex Darrow (alexd@contoso.com);Paul Walker (paulw@contoso.com);")
                                    done();
                                }
                            );




                        });

                    it("Get Cc recipients ",
                        function (done) {


                            /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                            /* Get Cc recipients (Applicable only on message) */
                            Office.context.mailbox.item.cc.getAsync(
                                function callback(asyncResult) {
                                    if (asyncResult.status == "failed") {
                                        console.log("Action failed with error: " + asyncResult.error.message);
                                        document.getElementById('getCcRecipients').innerHTML = ("Action failed with error: " + asyncResult.error.message);
                                    } else {
                                        var recipients = "";
                                        asyncResult.value.forEach(function (recipient, index) {
                                            recipients = recipients + recipient.displayName + " (" + recipient.emailAddress + ");";
                                        });
                                        console.log(recipients);
                                        document.getElementById('getCcRecipients').innerHTML = (recipients);
                                    }
                                    expect(asyncResult.status).toBe("succeeded");
                                    expect(recipients).toBe("Allie Bellew (allieb@contoso.com);Alex Darrow (alexd@contoso.com);Paul Walker (paulw@contoso.com);")
                                    done();
                                }
                            );




                        });


                   





                it("Get body content async",
                    function (done) {

                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get body content */
                        Office.context.mailbox.item.body.getAsync("text",
                            function (asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                } else {
                                    console.log(asyncResult.value);
                                    document.getElementById("messageBody").innerHTML = asyncResult.value;
                                }
                                expect(asyncResult.status).toBe("succeeded");
                                done();
                            }
                        );





                    });




             



        });
       
      
       



    });

