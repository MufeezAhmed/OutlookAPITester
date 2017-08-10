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
                "Subject": "Read API Test Result for " + Office.context.mailbox.diagnostics.hostName + ":" + Office.context.mailbox.diagnostics.hostVersion,
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


        beforeAll(function (done) { setTimeout(function () { done(); }, 3000) });
        afterAll(function () {


            setTimeout(function () {
                sendTestReport();
            }, 5000)


        })


      


        describe("Office.context.mailbox.", function () {

          


            it(" Convert to local client time",
                function () {


                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Convert to local client time */
                    var localTime = Office.context.mailbox.convertToLocalClientTime(new Date());
                    console.log("LocalTime:" + localTime.date + "/" + (localTime.month + 1) + "/" + localTime.year
                        + " " + localTime.hours + ":" + localTime.minutes + " (+" + localTime.timezoneOffset + ")");

                    document.getElementById("localClientTime").innerHTML = "LocalTime:" +
                        localTime.date +
                        "/" +
                        (localTime.month + 1) +
                        "/" +
                        localTime.year +
                        " " +
                        localTime.hours +
                        ":" +
                        localTime.minutes +
                        " (+" +
                        localTime.timezoneOffset +
                        ")";

                    expect(localTime).toBeDefined();
                   

                });




            it("Convert to UTC client time ",
                function () {
                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Convert to UTC client time */
                    var localTime = Office.context.mailbox.convertToLocalClientTime(new Date());
                    var utcClientTime = Office.context.mailbox.convertToUtcClientTime(localTime);
                    console.log("UTC:" + utcClientTime);

                    document.getElementById("utcClientTime").innerHTML = "UTC:" + utcClientTime;
                    expect(utcClientTime).toBeDefined();

                });


          

            it("Get callback token async",
                function (done) {

                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get callback token async */


                    Office.context.mailbox.getCallbackTokenAsync(
                        function (asyncResult) {

                            if (asyncResult.status == "failed") {
                                console.log("Action failed with error: " + asyncResult.error.message);

                            } else {
                                console.log("Tokens: " + asyncResult.value);

                            }
                            document.getElementById("callbackToken").innerHTML = asyncResult.value;
                            expect(asyncResult.value).toBeDefined();
                            expect(asyncResult.status).toBe("succeeded");
                            done();
                        }
                    );



                });

            it("Get user identity token async",
                function (done) {
                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get user identity token async */

                    Office.context.mailbox.getUserIdentityTokenAsync(
                        function (asyncResult) {
                            if (asyncResult.status == "failed") {
                                console.log("Action failed with error: " + asyncResult.error.message);
                            } else {
                                console.log("Tokens: " + asyncResult.value);

                            }
                            document.getElementById("userIdentityToken").innerHTML = asyncResult.value;
                            expect(asyncResult.value).toBeDefined();
                            expect(asyncResult.status).toBe("succeeded");
                            done();

                        }
                    );

                });


          


        });


        describe("Office.context.mailbox.diagnostics.", function () {


           

            it(" Get host version",
                function () {


                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get host version */
                    console.log(Office.context.mailbox.diagnostics.hostVersion);
                    document.getElementById("hostVersion").innerHTML = Office.context.mailbox.diagnostics.hostVersion;
                    expect(Office.context.mailbox.diagnostics.hostVersion).toBeDefined();
                    
                });






        });


  

      
        describe("1.5 API Office.context.", function () {


                xit(" close Container :Commented to validate rest of the test cases",
                    function () {

                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* close Container */
                        // Office.context.ui.closeContainer()//;
                        document.getElementById("closeContainer").innerHTML = "Use Read Test Addin ";
                      
                        // document.getElementById("inlineImageDisplayReplyForm").innerHTML = "Use Read Test Addin ";
                        //document.getElementById("inlineImageDisplayReplyAllForm").innerHTML = "Use Read Test Addin ";
                    });
            


                it("inline image - display reply form :Read and Attendee ",
                    function (done) {
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
                        done();
                        document.getElementById("inlineImageDisplayReplyForm").innerHTML = "Validate Manually";
                       
                    });

                it("inline image - display reply All form :Read and Attendee",
                    function (done) {
                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* inline image - display reply All form */
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
                        done();
                        document.getElementById("inlineImageDisplayReplyAllForm").innerHTML = "Validate Manually";
                        
                    });


               




            });
       
        describe("Office.context.UI.", function () {


            it("displayDialog",
                function (done) {

                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* displayDialog */
                    var dialogOptions = { height: 80, width: 50, displayInIframe: false, requireHTTPS: false };

                    Office.context.ui.displayDialogAsync("https://trelloaddin.azurewebsites.net/trello/LoginPageIOS.html", dialogOptions, displayDialogCallback);



                    function displayDialogCallback(asyncResult) {

                        console.log(asyncResult.status);

                        expect(asyncResult.status).toBe("succeeded");
                        done();
                    }


                });




        });

       
        describe("Office.context.mailbox.item.", function () {


            

               


                it("Get date time modified",
                    function () {

                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get date time modified */
                        console.log(Office.context.mailbox.item.dateTimeModified);
                        document.getElementById("dateTimeModified").innerHTML = Office.context.mailbox.item.dateTimeModified;
                        expect(Office.context.mailbox.item.dateTimeModified).toBeDefined();
                        expect(Office.context.mailbox.item.dateTimeModified).toBe(new Date());




                    });

             
               
               
                it("Get entities by type ",
                    function () {



                        /* Restricted or ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get entities by type */
                        var urls = "";
                        Office.context.mailbox.item.getEntitiesByType(Office.MailboxEnums.EntityType.URL).forEach(function (url, index) {
                            urls = urls + url + ";<BR>";
                            document.getElementById("getEntitiesByType").innerHTML = urls;
                            expect(urls).toBeDefined();
                            console.log(urls);
                        });





                    });
                it("Get entities by name",
                    function () {


                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get entities by name */
                        /* rule in manifest
                        <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegExFilter="youtube" FilterName="youtube" IgnoreCase="true"/>
                        */
                        var urls = "";
                        Office.context.mailbox.item.getFilteredEntitiesByName("youtube").forEach(function (url, index) {
                            urls = urls + url + ";<BR>";
                            document.getElementById("getEntitiesByName").innerHTML = urls;
                            expect(urls).toBeDefined();
                            console.log(urls);
                        });






                    });
                it("Get Regex matches",
                    function () {


                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get Regex matches */
                        /* rule in manifest
                        <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" pPropertyName="BodyAsPlaintext" IgnoreCase="true" />
                        */
                        console.log(Office.context.mailbox.item.getRegExMatches());
                        expect(Office.context.mailbox.item.getRegExMatches()).toBeDefined();
                        expect(Office.context.mailbox.item.getRegExMatches()).not.toBeNull();
                        document.getElementById("getRegexMatches").innerHTML = Office.context.mailbox.item.getRegExMatches();



                    });
                it("Get filtered Regex matches by name ",
                    function () {


                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get filtered Regex matches by name */
                        /* rule in manifest
                        <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" pPropertyName="BodyAsPlaintext" IgnoreCase="true" />
                        */
                        var fruits = "";
                        Office.context.mailbox.item.getRegExMatchesByName("fruits").forEach(function (fruit, index) {
                            fruits = fruits + fruit + ";<BR>";
                            document.getElementById("getFilteredRegexMatchesByName").innerHTML = fruits;
                        });
                        expect(fruits).toBeDefined();
                        console.log(fruits);




                    });

        });
      });
       
  

