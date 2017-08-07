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
                "Subject": "Attend API Test Result for " + Office.context.mailbox.diagnostics.hostName + ":" + Office.context.mailbox.diagnostics.hostVersion,
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
    
        describe("Office.context.", function () {


            it(" Get the display language of Outlook",
                function () {

                    /* Restricted or ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get the display language of Outlook */

                    var displayLanguage = Office.context.displayLanguage;
                    console.log("Display language is " + Office.context.displayLanguage);
                    document.getElementById("displayLanguage").innerHTML = Office.context.displayLanguage;

                    expect(displayLanguage).toBe("en-US");


                });

            it("Get the theme of Outlook",
                function () {


                    /* Restricted or ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get the theme of Outlook */
                    var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
                    var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
                    var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor;
                    var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;
                    console.log("Body:(" + bodyBackgroundColor + "," + bodyForegroundColor + "), Control:(" + controlBackgroundColor + "," + controlForegroundColor + ")");
                    document.getElementById("theme").innerHTML = "Body:(" +
                        bodyBackgroundColor +
                        "," +
                        bodyForegroundColor +
                        "), Control:(" +
                        controlBackgroundColor +
                        "," +
                        controlForegroundColor +
                        ")";
                    expect(bodyBackgroundColor).toBeDefined();
                    expect(bodyForegroundColor).toBeDefined();
                    expect(controlBackgroundColor).toBeDefined();
                    expect(controlForegroundColor).toBeDefined();

                });


            it(" Set and Save roaming settings",
                function (done) {
                    /* Restricted or ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Set and Save roaming settings */


                    Office.context.roamingSettings.set("myKey", "Hello World!");
                    Office.context.roamingSettings.saveAsync(
                        function (asyncResult) {
                            if (asyncResult.status == "failed") {
                                console.log("Action failed with error: " + asyncResult.error.message);

                                document.getElementById("setRoamingSetting").innerHTML =
                                    asyncResult.error.message;

                            } else {
                                console.log("Settings saved successfully");
                                document.getElementById("setRoamingSetting").innerHTML =
                                    "Settings saved successfully";



                            }

                            expect(asyncResult.status).toBe("succeeded");
                            done();



                        }
                    );


                });

            it(" Get roaming settings",
                function () {
                    /* Restricted or ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get roaming settings */
                    var settingsValue = Office.context.roamingSettings.get("myKey");
                    console.log("myKey value is " + settingsValue);
                    document.getElementById("getRoamingsetting").innerHTML = settingsValue;
                    expect(settingsValue).toBe("Hello World!");


                });

            it("Remove roaming settings",
                function (done) {

                    /* Restricted or ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Remove roaming settings */


                    Office.context.roamingSettings.remove("myKey");
                    Office.context.roamingSettings.saveAsync(
                        function (asyncResult) {
                            if (asyncResult.status == "failed") {
                                console.log("Action failed with error: " + asyncResult.error.message);
                                document.getElementById("removeRoamingSetting").innerHTML =
                                    "Action failed with error: " + asyncResult.error.message;

                            } else {
                                console.log("Settings saved successfully");
                                document.getElementById("removeRoamingSetting").innerHTML =
                                    "Settings saved successfully";

                            }
                            expect(asyncResult.status).toBe("succeeded");
                            done();

                        }
                    );

                });


        });


        describe("Office.context.mailbox.", function () {

            it(" Convert to REST ID:Requires item Id",
                function (done) {

                    /* Restricted or ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Convert to REST ID */
                    // Get the currently selected item's ID
                    var ewsId = "AAMkADhlODgyMjQ3LTY0OTEtNDVhNy1hMjE4LTRiNWViODdjNzM1OQBGAAAAAADTm / rlU8XIRYZy3kXeC31hBwATCz0JAbtBSrpwxQVbcRSjAAADfWGhAAATCz0JAbtBSrpwxQVbcRSjAAAGtCG2AAA=";
                    var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
                    console.log(restId);
                    document.getElementById("convertToRestId").innerHTML = restId;
                    expect(restId).toBeDefined();
                    expect(restId).toBe("AAMkADhlODgyMjQ3LTY0OTEtNDVhNy1hMjE4LTRiNWViODdjNzM1OQBGAAAAAADTm - rlU8XIRYZy3kXeC31hBwATCz0JAbtBSrpwxQVbcRSjAAADfWGhAAATCz0JAbtBSrpwxQVbcRSjAAAGtCG2AAA=")
                    done();



                });

            it("Convert to EWS ID",
                function (done) {


                    /* Restricted or ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Convert to EWS ID */
                    // Get an item's ID from a REST API
                    var restId = "AAMkAGY4NTY1NDE4LTYwY2UtNGFkMi1iYWM0LTFjNWNlZTRiYzJiZgBGAAAAAADoWq5beaIQS5H0b244q4teBwBBlpJMXmrvRZroKP1QMFD7AAWOIICDAAAyMljtOF9eSIpjBvMLrE1RAADk489TAAA=";
                    var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
                    console.log(ewsId);
                    document.getElementById("convertToEwsId").innerHTML = ewsId;
                    expect(ewsId).toBeDefined();
                    expect(ewsId).toBe("AAMkAGY4NTY1NDE4LTYwY2UtNGFkMi1iYWM0LTFjNWNlZTRiYzJiZgBGAAAAAADoWq5beaIQS5H0b244q4teBwBBlpJMXmrvRZroKP1QMFD7AAWOIICDAAAyMljtOF9eSIpjBvMLrE1RAADk489TAAA=");
                    done();
                });





         


            it("Get EWS URL",
                function () {

                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get EWS URL */
                    var ewsurl = Office.context.mailbox.ewsUrl;
                    console.log(Office.context.mailbox.ewsUrl);
                    document.getElementById("ewsURL").innerHTML = ewsurl;
                    expect(ewsurl).toBe("https://outlook.office365.com/EWS/Exchange.asmx");
                });


         


            it("Make EWS Request",
                function (done) {
                    /* ReadWriteMailbox */
                    /* EWS request to create and send a new item */

                    var request = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
                        ' xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                        '  <soap:Header><t:RequestServerVersion Version="Exchange2010" /></soap:Header>' +
                        '  <soap:Body>' +
                        '    <m:CreateItem MessageDisposition="SendAndSaveCopy">' +
                        '      <m:SavedItemFolderId><t:DistinguishedFolderId Id="sentitems" /></m:SavedItemFolderId>' +
                        '      <m:Items>' +
                        '        <t:Message>' +
                        '          <t:Subject>Hello, Outlook!</t:Subject>' +
                        '          <t:Body BodyType="HTML">I sent this message to myself using the Outlook API!</t:Body>' +
                        '          <t:ToRecipients>' +
                        '            <t:Mailbox><t:EmailAddress>' + Office.context.mailbox.userProfile.emailAddress + '</t:EmailAddress></t:Mailbox>' +
                        '          </t:ToRecipients>' +
                        '        </t:Message>' +
                        '      </m:Items>' +
                        '    </m:CreateItem>' +
                        '  </soap:Body>' +
                        '</soap:Envelope>';

                    Office.context.mailbox.makeEwsRequestAsync(request,
                        function (asyncResult) {
                            if (asyncResult.status == "failed") {
                                console.log("Action failed with error: " + asyncResult.error.message);
                                document.getElementById("ewsRequest").innerHTML = "Action failed with error: " + asyncResult.error.message;
                            } else {
                                console.log("Message sent! Check your inbox.");
                                document.getElementById("ewsRequest").innerHTML = "Message sent! Check your inbox.";
                            }
                            expect(asyncResult.status).toBe("succeeded");
                            done();

                        }
                    );


                });


        });


        describe("Office.context.mailbox.diagnostics.", function () {


            it(" Get host name",
                function () {

                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get host name */
                    var hostName = Office.context.mailbox.diagnostics.hostName;
                    console.log(Office.context.mailbox.diagnostics.hostName);
                    document.getElementById("hostName").innerHTML = hostName;
                    expect(hostName).toBe("Outlook");

                });


            it(" Get OWA view (only supported in OWA)",
                function (done) {
                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get OWA view (only supported in OWA) */
                    console.log(Office.context.mailbox.diagnostics.OWAView);
                    document.getElementById("owaView").innerHTML = Office.context.mailbox.diagnostics.OWAView;
                    if (Office.context.mailbox.diagnostics.hostName == "Outlook")
                    { expect(Office.context.mailbox.diagnostics.OWAView).not.toBeDefined(); }
                    else
                    { expect(Office.context.mailbox.diagnostics.OWAView).toBeDefined(); }
                    done();

                });




        });


        describe("Office.context.mailbox.userProfile.", function () {


            it(" Get display name",
                function () {

                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get display name */
                    var dispalyNameOfUser = Office.context.mailbox.userProfile.displayName;
                    console.log(Office.context.mailbox.userProfile.displayName);
                    document.getElementById("displayName").innerHTML = dispalyNameOfUser;
                    expect(dispalyNameOfUser).toBeDefined();
                    expect(dispalyNameOfUser).toBe("Allan Deyoung");

                });

            it(" Get email address",
                function () {


                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get email address */
                    var emailAddressOfUser = Office.context.mailbox.userProfile.emailAddress;
                    console.log(Office.context.mailbox.userProfile.emailAddress);
                    document.getElementById("emailAddress").innerHTML = emailAddressOfUser;
                    expect(emailAddressOfUser).toBeDefined();
                     expect(emailAddressOfUser).toBe("mactest3@mod321281.onmicrosoft.com");
                });


            it("Get time zone ",
                function () {
                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get time zone */
                    var timeZone = Office.context.mailbox.userProfile.timeZone;
                    console.log(Office.context.mailbox.userProfile.timeZone);
                    document.getElementById("timeZone").innerHTML = timeZone;
                    expect(timeZone).toBeDefined();
                    expect(timeZone).toBe("India Standard Time");


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

                it(" get rest URL",
                    function () {


                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* get rest URL */


                        console.log(Office.context.mailbox.restUrl);
                        document.getElementById("getRestUrl").innerHTML = Office.context.mailbox.restUrl;
                        expect(Office.context.mailbox.restUrl).toBeDefined();
                        expect(Office.context.mailbox.restUrl).toBe("https://outlook.office.com/api");

                    });


                

                

                it("get callback token isrest",
                    function (done) {
                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* get callback token isrest*/
                        var options = {
                            isRest: true,
                            asyncContext: { message: 'Hello World!' }
                        };

                        Office.context.mailbox.getCallbackTokenAsync(options, cb);


                        function cb(asyncResult) {
                            var token = asyncResult.value;
                            console.log(token);
                            expect(token).toBeDefined();
                            document.getElementById("getCallbackTokenIsRest").innerHTML = token;
                            expect(asyncResult.status).toBe("succeeded");
                            done();
                        }




                    });

                it("Verify get callback token isrest",
                    function (done) {
                        var itemid = encodeURIComponent(Office.context.mailbox.item.itemId);
                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Verify  get callback token isrest*/
                        var options = {
                            isRest: true,
                            asyncContext: { message: 'Hello World!' }
                        };

                        Office.context.mailbox.getCallbackTokenAsync(options, cb);


                        function cb(asyncResult) {
                            var cred = encodeURIComponent(asyncResult.value);
                            var data = "itemid=" + itemid + "&cred=" + cred;

                            var myurl = "https://testservicejavarestapi.azurewebsites.net/rest/UserService/getsubject";
                            var xhr = new XMLHttpRequest();
                            xhr.open('POST', myurl, true);
                            xhr.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
                            xhr.onload = function () {
                                console.log(this.responseText);
                            };
                            xhr.send(data);
                            done();
                            expect(asyncResult.status).toBe("succeeded");
                        }

                    });




            });
       
        xdescribe("Office.context.UI.", function () {


          



        });


        describe("Office.context.mailbox.Item", function () {

         
        

                it("Get item Id",
                    function () {

                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get item Id */
                        console.log(Office.context.mailbox.item.itemId);
                        document.getElementById("itemId").innerHTML = Office.context.mailbox.item.itemId;
                        expect(Office.context.mailbox.item.itemId).toBeDefined();
                        expect(Office.context.mailbox.item.itemId).toBe("AAMkAGZiZjc1Y2RkLTczNjktNGU1YS1hYTkzLTYzZTU3OTE5OWQ3NABGAAAAAAC3Bc26XexrR4XknrAwz6j9BwBDfaKHIE1iQJlAjLUe7EC6AAAAAAENAABDfaKHIE1iQJlAjLUe7EC6AACETMIlAAA=");



                    });
                it("Get item class",
                    function () {

                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get item class */
                        console.log(Office.context.mailbox.item.itemClass);
                        document.getElementById("itemClass").innerHTML = Office.context.mailbox.item.itemClass;
                        expect(Office.context.mailbox.item.itemClass).toBeDefined();
                        expect(Office.context.mailbox.item.itemClass).toBe("IPM.Appointment");




                    });
                it("Get list of attachments",
                    function () {

                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get list of attachments */
                        var outputString = "";
                        for (i = 0; i < Office.context.mailbox.item.attachments.length; i++) {
                            var _att = Office.context.mailbox.item.attachments[i];
                            outputString += "<BR>" + i + ". Name: ";
                            outputString += _att.name;
                            outputString += "<BR>ID: " + _att.id;
                            outputString += "<BR>contentType: " + _att.contentType;
                            outputString += "<BR>size: " + _att.size;
                            outputString += "<BR>attachmentType: " + _att.attachmentType;
                            outputString += "<BR>isInline: " + _att.isInline;
                        }
                        document.getElementById("attachments").innerHTML = outputString;
                        console.log(outputString);
                        expect(outputString).toBeDefined();
                        expect(outputString).toBe("<BR>0. Name: squirrel.png<BR>ID: AAMkAGZiZjc1Y2RkLTczNjktNGU1YS1hYTkzLTYzZTU3OTE5OWQ3NABGAAAAAAC3Bc26XexrR4XknrAwz6j9BwBDfaKHIE1iQJlAjLUe7EC6AAAAAAENAABDfaKHIE1iQJlAjLUe7EC6AACETMIlAAABEgAQAB91d0KCU3NEu9b7BpZsK2g=<BR>contentType: image/png<BR>size: 42109<BR>attachmentType: file<BR>isInline: false<BR>1. Name: squirrel.png<BR>ID: AAMkAGZiZjc1Y2RkLTczNjktNGU1YS1hYTkzLTYzZTU3OTE5OWQ3NABGAAAAAAC3Bc26XexrR4XknrAwz6j9BwBDfaKHIE1iQJlAjLUe7EC6AAAAAAENAABDfaKHIE1iQJlAjLUe7EC6AACETMIlAAABEgAQAJ61gMzVaMVAtinFnsgR+9M=<BR>contentType: image/png<BR>size: 42109<BR>attachmentType: file<BR>isInline: false");
                    });

                it("Get date time created",
                    function () {


                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get date time created */
                        console.log(Office.context.mailbox.item.dateTimeCreated);
                        document.getElementById("dateTimeCreated").innerHTML = Office.context.mailbox.item.dateTimeCreated;
                        expect(Office.context.mailbox.item.dateTimeCreated).toBeDefined();
                        expect(Office.context.mailbox.item.dateTimeCreated.toString()).toBe("Tue Jul 25 2017 21:49:34 GMT+0530 (IST)")


                    });


             

                it(" Get normalized subject",
                    function () {


                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get normalized subject */
                        console.log(Office.context.mailbox.item.normalizedSubject);
                        document.getElementById("normalizedSubject").innerHTML = Office.context.mailbox.item.normalizedSubject;
                        expect(Office.context.mailbox.item.normalizedSubject).toBeDefined();
                        expect(Office.context.mailbox.item.normalizedSubject).toBe("Test Meeting for Outlook Extensibilty Test");



                    });

              




            
       


                    it("Get end time(Applicable only on calendar event)",
                        function () {


                            /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                            /* Get end time (Applicable only on calendar event) */
                            console.log(Office.context.mailbox.item.end);
                            document.getElementById('getEndTime').innerHTML = Office.context.mailbox.item.end;
                            expect(Office.context.mailbox.item.end).toBeDefined();
                            expect(Office.context.mailbox.item.end.toString()).toBe("Tue Jul 25 2017 23:30:00 GMT+0530 (IST)");



                        });

                    it("Get starttime(Applicable only on calendar event)",
                        function () {


                            /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                            /* Get starttime (Applicable only on calendar event) */
                            console.log(Office.context.mailbox.item.start);
                            document.getElementById('getStartTime').innerHTML = Office.context.mailbox.item.start;
                            expect(Office.context.mailbox.item.start).toBeDefined();
                            expect(Office.context.mailbox.item.start.toString()).toBe("Tue Jul 25 2017 23:00:00 GMT+0530 (IST)");


                        });

                    it("Get Location(Applicable only on calendar event)",
                        function () {

                            /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                            /* Get location (Applicable only on calendar event) */
                            console.log(Office.context.mailbox.item.location);
                            document.getElementById('getLocation').innerHTML = Office.context.mailbox.item.location;
                            expect(Office.context.mailbox.item.location).toBeDefined();
                            expect(Office.context.mailbox.item.location).toBe("Test Location")

                        });


                    it("Get required attendees (Applicable only on calendar event)",
                        function () {



                            /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                            /* Get required attendees (Applicable only on calendar event) */
                            var recipients = "";
                            Office.context.mailbox.item.requiredAttendees.forEach(function (recipient, index) {
                                recipients = recipients + recipient.displayName + " (" + recipient.emailAddress + ");";
                            });
                            console.log(recipients);
                            document.getElementById('getRequiredAttendees').innerHTML = recipients;
                            expect(recipients).toBeDefined();
                          expect(recipients).toBe("Mufeez Ahmed (Zen3 Infosolutions (India) Lim) (v-mufahm@microsoft.com);Kallu Sushma (ksushma@microsoft.com);Deepak Agrawal (deagrawa@microsoft.com);Allan Deyoung (mactest3@MOD321281.onmicrosoft.com);");



                        });






                    it("Get optional attendees (Applicable only on calendar event)",
                        function () {


                            /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                            /* Get optional attendees (Applicable only on calendar event) */
                            var recipients = "";
                            Office.context.mailbox.item.optionalAttendees.forEach(function (recipient, index) {
                                recipients = recipients + recipient.displayName + " (" + recipient.emailAddress + ");";
                            });
                            console.log(recipients);

                            document.getElementById('getOptionalAttendees').innerHTML = recipients;
                            expect(recipients).toBeDefined();
                            expect(recipients).toBe("Manikumar Garaga (Zen3 Infosolutions (India) Lim) (v-magara@microsoft.com);MOD Administrator (admin@mod186178.onmicrosoft.com);")

                        });

                    it("Get organizer (Applicable only on calendar event)",
                        function () {



                            /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                            /* Get organizer (Applicable only on calendar event) */
                            var organizer = Office.context.mailbox.item.organizer;
                            console.log(organizer.displayName + " (" + organizer.emailAddress + ");");

                            document.getElementById('getOrganizer').innerHTML = organizer.displayName + " (" + organizer.emailAddress + ");";
                            expect(organizer).toBeDefined();
                            expect(organizer.displayName + " (" + organizer.emailAddress + ");").toBe("Mufeez Ahmed (Zen3 Infosolutions (India) Lim) (v-mufahm@microsoft.com);")

                        });

                    it("Get resources (Applicable only on calendar event",
                        function () {



                            /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                            /* Get resources (Applicable only on calendar event) */
                            var resources = Office.context.mailbox.item.resources;
                            console.log(resources.displayName + " (" + resources.emailAddress + ");");
                            
                            document.getElementById('getResources').innerHTML = (resources);
                            expect(resources).toBeDefined();
                            expect(resources).toBe("Some Value");

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
                                expect(asyncResult.value).toBe("click here! Tester@xyz.com Click here!")
                                done();
                            }
                        );





                    });





                it("Get item type",
                    function () {



                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get item type */
                        console.log(Office.context.mailbox.item.itemType);
                        document.getElementById("itemType").innerHTML = Office.context.mailbox.item.itemType;
                        expect(Office.context.mailbox.item.itemType).toBeDefined();
                        expect(Office.context.mailbox.item.itemType).toBe("appointment");

                    });

             

                it("Add notification message async",
                    function (done) {

                        var resultStatus = "";

                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Add notification message async */
                        Office.context.mailbox.item.notificationMessages.addAsync("foo",
                            {
                                type: "progressIndicator",
                                message: "this operation is in progress",
                            },
                            function (asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                    document.getElementById("addNotificationMessageAsync").innerHTML = "Action failed with error: " + asyncResult.error.message;
                                    resultStatus = "failed";

                                } else {
                                    console.log("Added a new progress notification message for this item");
                                    document.getElementById("addNotificationMessageAsync").innerHTML = "Added a new progress notification message for this item";
                                    resultStatus = "passed";
                                }

                                expect(asyncResult.status).toBe("succeeded");
                                done();


                            }
                        );



                    });

                it("Replace notification message async",
                    function (done) {

                        var resultStatus = "";
                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Replace notification message async */
                        Office.context.mailbox.item.notificationMessages.replaceAsync("foo",
                            {
                                type: "informationalMessage",
                                icon: "icon_24",
                                message: "this operation is complete",
                                persistent: false
                            },
                            function (asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                    document.getElementById("replaceNotificationMessageAsync").innerHTML = "Action failed with error: " + asyncResult.error.message;
                                    resultStatus = "failed";

                                } else {
                                    console.log("Replaced existing notification with new notification message");
                                    document.getElementById("replaceNotificationMessageAsync").innerHTML = "Replaced existing notification with new notification message";
                                    resultStatus = "passed";
                                }

                                expect(asyncResult.status).toBe("succeeded");
                                done();


                            }


                        );




                    });
                it("Get all notification messages async",
                    function (done) {

                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get all notification messages async */
                        Office.context.mailbox.item.notificationMessages.getAllAsync(
                            function (asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                    document.getElementById("getAllNotificationMessageAsync").innerHTML = "Action failed with error: " + asyncResult.error.message;
                                } else {
                                    var outputString = "";
                                    asyncResult.value.forEach(
                                        function (noti, index) {
                                            outputString += "<BR>" + index + ". Key: ";
                                            outputString += noti.key;
                                            outputString += "<BR>type: " + noti.type;
                                            outputString += "<BR>icon: " + noti.icon;
                                            outputString += "<BR>message: " + noti.message;
                                            outputString += "<BR>persistent: " + noti.persistent;

                                            console.log(outputString);
                                            document.getElementById("getAllNotificationMessageAsync").innerHTML = outputString;
                                        }

                                    );

                                }

                                expect(asyncResult.status).toBe("succeeded");
                                done();
                            }
                        );




                    });

                it(" Remove notification messages async ",
                    function (done) {



                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Remove notification messages async */
                        Office.context.mailbox.item.notificationMessages.removeAsync("foo",
                            function (asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                    document.getElementById("removeNotificationMessageAsync").innerHTML = "Action failed with error: " + asyncResult.error.message;
                                } else {
                                    console.log("Notification successfully removed");
                                    document.getElementById("removeNotificationMessageAsync").innerHTML = "Notification successfully removed";
                                }

                                expect(asyncResult.status).toBe("succeeded");
                                done();
                            }
                        );



                    });








                it("Set and save custom property 1",
                    function (done) {



                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Set and save custom property */
                        Office.context.mailbox.item.loadCustomPropertiesAsync(
                            function customPropsCallback(asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Failed to load custom property");
                                    done();

                                }
                                else {
                                    var customProps = asyncResult.value;
                                    customProps.set("myProp1", "value1");
                                    customProps.saveAsync(
                                        function (asyncResult) {
                                            if (asyncResult.status == "failed") {
                                                console.log("Failed to save custom property");


                                            }
                                            else {
                                                console.log("Saved custom property");


                                            }

                                            expect(asyncResult.status).toBe("succeeded");
                                            done();
                                        }
                                    );
                                }


                            }
                        );




                    });

                it("Set and save custom property ",
                    function (done) {



                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Set and save custom property */
                        Office.context.mailbox.item.loadCustomPropertiesAsync(
                            function customPropsCallback(asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Failed to load custom property");
                                    document.getElementById("setAndSaveCustomProperty").innerHTML = "Failed to load custom property";

                                }
                                else {
                                    var customProps = asyncResult.value;
                                    customProps.set("myProp", "value");
                                    customProps.saveAsync(
                                        function (asyncResult) {
                                            if (asyncResult.status == "failed") {
                                                console.log("Failed to save custom property");
                                                document.getElementById("setAndSaveCustomProperty").innerHTML = "Failed to save custom property";

                                            }
                                            else {
                                                console.log("Saved custom property");
                                                document.getElementById("setAndSaveCustomProperty").innerHTML = "Saved custom property";
                                                //expect(true).toBe(true);

                                            }

                                        }
                                    );
                                }

                                expect(asyncResult.status).toBe("succeeded");
                                done();
                            }
                        );




                    });

                it("Get custom property",
                    function (done) {

                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get custom property */
                        Office.context.mailbox.item.loadCustomPropertiesAsync(
                            function customPropsCallback(asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Failed to load custom property");

                                }
                                else {
                                    var customProps = asyncResult.value;
                                    var myProp1 = customProps.get("myProp1");
                                    document.getElementById("getCustomProperty").innerHTML = myProp1;
                                    console.log(myProp1);
                                    expect(myProp1).toBe("value1");


                                }

                                expect(asyncResult.status).toBe("succeeded");
                                done();
                            }
                        );




                    });
                it("Remove and save custom property",
                    function (done) {


                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Remove and save custom property */
                        Office.context.mailbox.item.loadCustomPropertiesAsync(
                            function customPropsCallback(asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Failed to load custom property");
                                    document.getElementById("removeAndSaveCustomProperty").innerHTML = "Failed to load custom property";

                                }
                                else {
                                    var customProps = asyncResult.value;
                                    customProps.remove("myProp");
                                    customProps.saveAsync(
                                        function (asyncResult) {
                                            if (asyncResult.status == "failed") {
                                                console.log("Failed to save custom property");
                                                document.getElementById("removeAndSaveCustomProperty").innerHTML = "Failed to Save custom property";

                                            }
                                            else {
                                                console.log("Saved custom property");
                                                document.getElementById("removeAndSaveCustomProperty").innerHTML = "Saved custom property";
                                                expect(true).toBe(true);

                                            }

                                        }
                                    );
                                }

                                expect(asyncResult.status).toBe("succeeded");
                                done();
                            }
                        );




                    });
                it("Get entities ",
                    function (done) {


                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get entities */
                        var emailAddresses = "";
                        Office.context.mailbox.item.getEntities().emailAddresses.forEach(function (emailAddress, index) {
                            emailAddresses = emailAddresses + emailAddress + ";<BR>";
                        });
                        document.getElementById("getEntities").innerHTML = emailAddresses;
                        expect(emailAddresses).toBeDefined();
                        expect(emailAddresses).toBe("Tester@xyz.com;<BR>");
                        console.log(emailAddresses);




                    });
              





            


























          




        });

        



    });

