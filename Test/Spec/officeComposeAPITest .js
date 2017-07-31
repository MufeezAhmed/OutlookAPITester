﻿Office.initialize = function (reason) { };

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
        describe("Office.context", function () {


            it(" Get the display language of Outlook",
                function (done) {

                    /* Restricted or ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get the display language of Outlook */

                    var displayLanguage = Office.context.displayLanguage;
                    console.log("Display language is " + Office.context.displayLanguage);
                    document.getElementById("displayLanguage").innerHTML = Office.context.displayLanguage;

                    expect(displayLanguage).toBe("en-US");
                    done();

                });

            it("Get the theme of Outlook",
                function (done) {


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
                    done();
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
                function (done) {
                    /* Restricted or ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get roaming settings */
                    var settingsValue = Office.context.roamingSettings.get("myKey");
                    console.log("myKey value is " + settingsValue);
                    document.getElementById("getRoamingsetting").innerHTML = settingsValue;
                    expect(settingsValue).toBe("Hello World!");
                    done();

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


        describe("Office.context.mailbox", function () {


            it(" Convert to REST ID:Requires ITEM Id",
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


            it(" Convert to local client time",
                function (done) {


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
                    done();

                });




            it("Convert to UTC client time ",
                function (done) {
                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Convert to UTC client time */
                    var localTime = Office.context.mailbox.convertToLocalClientTime(new Date());
                    var utcClientTime = Office.context.mailbox.convertToUtcClientTime(localTime);
                    console.log("UTC:" + utcClientTime);

                    document.getElementById("utcClientTime").innerHTML = "UTC:" + utcClientTime;
                    expect(utcClientTime).toBeDefined();
                    done();
                });


            it("Get EWS URL",
                function (done) {

                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get EWS URL */
                    var ewsurl = Office.context.mailbox.ewsUrl;
                    console.log(Office.context.mailbox.ewsUrl);
                    document.getElementById("ewsURL").innerHTML = ewsurl;
                    expect(ewsurl).toBe("https://outlook.office365.com/EWS/Exchange.asmx");
                    done();
                });


            it("Get callback token async",
                function (done) {
                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get callback token async */
                    Office.context.mailbox.getCallbackTokenAsync(
                        function (asyncResult) {
                            if (asyncResult.status == "failed") {
                                console.log("Action failed with error: " + asyncResult.error.message);
                                document.getElementById("callbackToken").innerHTML = "Action failed with error: " + asyncResult.error.message;
                            } else {
                                console.log("Tokens: " + asyncResult.value);
                                document.getElementById("callbackToken").innerHTML = "Tokens: " + asyncResult.value;
                            }
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


            it("Make EWS Request",
                function (done) {
                    /* ReadWriteMailbox */
                    /* EWS request to create and send a new item */

                    var htmlcode = '<html><body><h1>My First Heading</h1><p>My first paragraph.</p></body></html>'

                    var request = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
                        ' xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                        '  <soap:Header><t:RequestServerVersion Version="Exchange2010" /></soap:Header>' +
                        '  <soap:Body>' +
                        '    <m:CreateItem MessageDisposition="SendAndSaveCopy">' +
                        '      <m:SavedItemFolderId><t:DistinguishedFolderId Id="sentitems" /></m:SavedItemFolderId>' +
                        '      <m:Items>' +
                        '        <t:Message>' +
                        '          <t:Subject>Hello, Outlook!</t:Subject>' +
                        '          <t:Body BodyType="HTML">&lt;a id="LPNoLP" href="http://www.contoso.com" &gt;Click here!&lt;/a&gt;</t:Body>' +
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


        describe("Office.context.mailbox.diagnostics", function () {


            it(" Get host name",
                function (done) {

                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get host name */
                    var hostName = Office.context.mailbox.diagnostics.hostName;
                    console.log(Office.context.mailbox.diagnostics.hostName);
                    document.getElementById("hostName").innerHTML = hostName;
                    expect(hostName).toBe("Outlook");
                    done();
                });

            it(" Get host version",
                function (done) {


                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get host version */
                    console.log(Office.context.mailbox.diagnostics.hostVersion);
                    document.getElementById("hostVersion").innerHTML = Office.context.mailbox.diagnostics.hostVersion;
                    expect(Office.context.mailbox.diagnostics.hostVersion).toBeDefined();
                    done();
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


        describe("Office.context.mailbox.userProfile", function () {


            it(" Get display name",
                function (done) {

                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get display name */
                    var dispalyNameOfUser = Office.context.mailbox.userProfile.displayName;
                    console.log(Office.context.mailbox.userProfile.displayName);
                    document.getElementById("displayName").innerHTML = dispalyNameOfUser;
                    expect(dispalyNameOfUser).toBeDefined();
                    done();
                });

            it(" Get email address",
                function (done) {


                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get email address */
                    var emailAddressOfUser = Office.context.mailbox.userProfile.emailAddress;
                    console.log(Office.context.mailbox.userProfile.emailAddress);
                    document.getElementById("emailAddress").innerHTML = emailAddressOfUser;
                    expect(emailAddressOfUser).toBeDefined();
                    done();
                });


            it("Get time zone ",
                function (done) {
                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get time zone */
                    var timeZone = Office.context.mailbox.userProfile.timeZone;
                    console.log(Office.context.mailbox.userProfile.timeZone);
                    document.getElementById("timeZone").innerHTML = timeZone;
                    expect(timeZone).toBeDefined();
                    done();

                });




        });


        describe("1.5 API Office.context ", function () {


            it(" close Container :Commented to validate rest of the test cases",
                function () {

                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* close Container */
                    // Office.context.ui.closeContainer()//;
                    document.getElementById("closeContainer").innerHTML = "Use Read Test Addin ";
                    expect(true).toBe(true);
                    // document.getElementById("inlineImageDisplayReplyForm").innerHTML = "Use Read Test Addin ";
                    //document.getElementById("inlineImageDisplayReplyAllForm").innerHTML = "Use Read Test Addin ";
                });

            it(" get rest URL",
                function (done) {


                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* get rest URL */


                    console.log(Office.context.mailbox.restUrl);
                    document.getElementById("getRestUrl").innerHTML = Office.context.mailbox.restUrl;
                    expect(Office.context.mailbox.restUrl).toBeDefined();
                    done();
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
                    var itemid = encodeURIComponent("AAMkADhlODgyMjQ3LTY0OTEtNDVhNy1hMjE4LTRiNWViODdjNzM1OQBGAAAAAADTm / rlU8XIRYZy3kXeC31hBwATCz0JAbtBSrpwxQVbcRSjAAADfWGhAAATCz0JAbtBSrpwxQVbcRSjAAAGtCG2AAA=");
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

        describe("Office.context.UI", function () {


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


        describe("Office.context.mailbox.item", function () {

           


                it(" Set subject Async",
                    function (done) {


                        /* ReadItem??? or ReadWriteItem or ReadWriteMailbox */
                        /* Set subject */
                        Office.context.mailbox.item.subject.setAsync("New subject!",
                            function (asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                    document.getElementById("setSubjectAsync").innerHTML = "Action failed with error: " + asyncResult.error.message;
                                } else {
                                    console.log("Subject set successfully");
                                    document.getElementById("setSubjectAsync").innerHTML = "Subject set successfully";
                                }
                                expect(asyncResult.status).toBe("succeeded");
                                done();
                            }
                        );




                    });
                it("Set body content Async",
                    function (done) {

                        /* ReadWriteItem or ReadWriteMailbox */
                        /* Set body content */
                        Office.context.mailbox.item.body.setAsync(
                            '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
                            { coercionType: "html" },
                            function (asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                    document.getElementById("setBodyAsync").innerHTML = "Action failed with error: " + asyncResult.error.message;
                                } else {
                                    console.log("Successfully set body text");
                                    document.getElementById("setBodyAsync").innerHTML = "Body set successfully"
                                }
                                expect(asyncResult.status).toBe("succeeded");
                                done();
                            }
                        );





                    });



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
                                done();
                            }
                        );




                    });

                it("Prepend body content Async",
                    function (done) {

                        /* ReadWriteItem or ReadWriteMailbox */
                        /* Prepend body content */
                        Office.context.mailbox.item.body.prependAsync(
                            '<a id="LPNoLP" href="http://www.contoso.com">click here! </a> <h1> Tester@xyz.com <h1>',
                            { coercionType: "html" },
                            function (asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                    document.getElementById("prependBodyAsync").innerHTML = "Action failed with error: " + asyncResult.error.message;
                                } else {
                                    console.log("Successfully prepended body text");
                                    document.getElementById("prependBodyAsync").innerHTML = "Successfully prepended body text"
                                }
                                expect(asyncResult.status).toBe("succeeded");
                                done();
                            }
                        );





                    });

                it("Add file attachment Async",
                    function (done) {

                        /* ReadWriteItem or ReadWriteMailbox */
                        /* Add file attachment */
                        var attachmentURL = "http://i.imgur.com/sRgTlGR.jpg";
                        Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, "squirrel.png",
                            function callback(asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                    document.getElementById("addFileAttchmentAsync").innerHTML = "Action failed with error: " + asyncResult.error.message;
                                } else {
                                    console.log("Attachment added with identifier:" + asyncResult.value);
                                    document.getElementById("addFileAttchmentAsync").innerHTML = "Attachment added with identifier:" + asyncResult.value
                                }
                                expect(asyncResult.status).toBe("succeeded");
                                done();
                            }
                        );





                    });

                it("Add item attachment Async",
                    function (done) {


                        /* ReadWriteItem or ReadWriteMailbox */
                        /* Add item attachment */
                        // Item ID of a mail item
                        var itemId = "AAMkADhlODgyMjQ3LTY0OTEtNDVhNy1hMjE4LTRiNWViODdjNzM1OQBGAAAAAADTm/rlU8XIRYZy3kXeC31hBwATCz0JAbtBSrpwxQVbcRSjAAADfWGhAAATCz0JAbtBSrpwxQVbcRSjAAAGtCG2AAA=";
                        Office.context.mailbox.item.addItemAttachmentAsync(itemId, "myitemattachment",
                            function callback(asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                    document.getElementById("addItemAttachmentAsync").innerHTML = "Subject set successfully"

                                } else {
                                    console.log("Attachment added with identifier:" + asyncResult.value);
                                    document.getElementById("addItemAttachmentAsync").innerHTML = "Attachment added with identifier:" + asyncResult.value;
                                }
                                expect(asyncResult.status).toBe("succeeded");
                                done();
                            }
                        );




                    });

                it("Save Form Async",
                    function (done) {


                        /* ReadWriteItem or ReadWriteMailbox */
                        // Save Form
                        Office.context.mailbox.item.saveAsync(
                            function callback(asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                    document.getElementById("saveAsync").innerHTML = "Action failed with error: " + asyncResult.error.message;
                                } else {
                                    console.log("Saved item with identifier:" + asyncResult.value);
                                    document.getElementById("saveAsync").innerHTML = "Saved item with identifier:" + asyncResult.value;
                                }
                                expect(asyncResult.status).toBe("succeeded");
                                done();
                            }
                        );




                    });
                it("Remove attachment Async",
                    function (done) {


                        /* ReadWriteItem or ReadWriteMailbox */
                        /* Remove attachment */
                        // identifier of an attachment
                        var attachmentId = "0";
                        Office.context.mailbox.item.removeAttachmentAsync(attachmentId,
                            function callback(asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                    document.getElementById("removeItemAttachmentAsync").innerHTML = "Action failed with error: " + asyncResult.error.message;
                                } else {
                                    console.log("Removed attachment with identifier:" + attachmentId);
                                    document.getElementById("removeItemAttachmentAsync").innerHTML = "Removed attachment with identifier:" + attachmentId;
                                }
                                expect(asyncResult.status).toBe("succeeded");
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
                                done();
                            }
                        );




                    });

                it("Add file attachment Async",
                    function (done) {

                        /* ReadWriteItem or ReadWriteMailbox */
                        /* Add file attachment */
                        var attachmentURL = "http://i.imgur.com/sRgTlGR.jpg";
                        Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, "squirrel.png",
                            function callback(asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                    document.getElementById("addFileAttchmentAsync").innerHTML = "Action failed with error: " + asyncResult.error.message;
                                } else {
                                    console.log("Attachment added with identifier:" + asyncResult.value);
                                    document.getElementById("addFileAttchmentAsync").innerHTML = "Attachment added with identifier:" + asyncResult.value
                                }
                                expect(asyncResult.status).toBe("succeeded");
                                done();
                            }
                        );





                    });




                    it("Set To recipients ",
                        function (done) {


                            /* ReadWriteItem or ReadWriteMailbox */
                            /* Set To recipients (Applicable only on message) */
                            var newRecipients = [
                                {
                                    "displayName": "Allie Bellew",
                                    "emailAddress": "allieb@contoso.com"
                                },
                                {
                                    "displayName": "Alex Darrow",
                                    "emailAddress": "alexd@contoso.com"
                                }
                            ];
                            Office.context.mailbox.item.to.setAsync(newRecipients,
                                function callback(asyncResult) {
                                    if (asyncResult.status == "failed") {
                                        console.log("Action failed with error: " + asyncResult.error.message);
                                        document.getElementById('setToRecipients').innerHTML = ("Action failed with error: " + asyncResult.error.message);
                                    } else {
                                        console.log("To set successfully");
                                        document.getElementById('setToRecipients').innerHTML = ("To set successfully");
                                    }
                                    expect(asyncResult.status).toBe("succeeded");
                                    done();
                                }
                            );




                        });

                    it("Set Cc recipients  ",
                        function (done) {


                            /* ReadWriteItem or ReadWriteMailbox */
                            /* Set Cc recipients (Applicable only on message) */
                            var newRecipients = [
                                {
                                    "displayName": "Allie Bellew",
                                    "emailAddress": "allieb@contoso.com"
                                },
                                {
                                    "displayName": "Alex Darrow",
                                    "emailAddress": "alexd@contoso.com"
                                }
                            ];
                            Office.context.mailbox.item.cc.setAsync(newRecipients,
                                function callback(asyncResult) {
                                    if (asyncResult.status == "failed") {
                                        console.log("Action failed with error: " + asyncResult.error.message);
                                        document.getElementById('setccRecipients').innerHTML = ("Action failed with error: " + asyncResult.error.message);
                                    } else {
                                        console.log("Cc set successfully");
                                        document.getElementById('setCcRecipients').innerHTML = ("Cc set successfully");
                                    }
                                    expect(asyncResult.status).toBe("succeeded");
                                    done();
                                }
                            );




                        });

                    it("Add To recipients ",
                        function (done) {

                            /* ReadWriteItem or ReadWriteMailbox */
                            /* Add To recipients (Applicable only on message) */
                            var newRecipients = [
                                {
                                    "displayName": "Paul Walker",
                                    "emailAddress": "paulw@contoso.com"
                                }
                            ];
                            Office.context.mailbox.item.to.addAsync(newRecipients,
                                function callback(asyncResult) {
                                    if (asyncResult.status == "failed") {
                                        console.log("Action failed with error: " + asyncResult.error.message);
                                        document.getElementById('addToRecipients').innerHTML = ("Action failed with error: " + asyncResult.error.message);
                                    } else {
                                        console.log("To add successfully");
                                        document.getElementById('addToRecipients').innerHTML = ("To added successfully");
                                    }
                                    expect(asyncResult.status).toBe("succeeded");
                                    done();
                                }
                            );





                        });
                    it(" Add Cc recipients ",
                        function (done) {

                            /* ReadWriteItem or ReadWriteMailbox */
                            /* Add Cc recipients (Applicable only on message) */
                            var newRecipients = [
                                {
                                    "displayName": "Paul Walker",
                                    "emailAddress": "paulw@contoso.com"
                                }
                            ];
                            Office.context.mailbox.item.cc.addAsync(newRecipients,
                                function callback(asyncResult) {
                                    if (asyncResult.status == "failed") {
                                        console.log("Action failed with error: " + asyncResult.error.message);
                                        document.getElementById('addCcRecipients').innerHTML = ("Action failed with error: " + asyncResult.error.message);
                                    } else {
                                        console.log("Cc add successfully");
                                        document.getElementById('addCcRecipients').innerHTML = ("Cc added successfully");
                                    }
                                    expect(asyncResult.status).toBe("succeeded");
                                    done();
                                }
                            );





                        });

                    it("Set Bcc recipients ",
                        function (done) {

                            /* ReadWriteItem or ReadWriteMailbox */
                            /* Set Bcc recipients (Applicable only on message) */
                            var newRecipients = [
                                {
                                    "displayName": "Allie Bellew",
                                    "emailAddress": "allieb@contoso.com"
                                },
                                {
                                    "displayName": "Alex Darrow",
                                    "emailAddress": "alexd@contoso.com"
                                }
                            ];
                            Office.context.mailbox.item.bcc.setAsync(newRecipients,
                                function callback(asyncResult) {
                                    if (asyncResult.status == "failed") {
                                        console.log("Action failed with error: " + asyncResult.error.message);
                                        document.getElementById('setBccRecipients').innerHTML = ("Action failed with error: " + asyncResult.error.message);
                                    } else {
                                        console.log("Bcc set successfully");
                                        document.getElementById('setBccRecipients').innerHTML = ("Bcc set successfully");
                                    }
                                    expect(asyncResult.status).toBe("succeeded");
                                    done();
                                }
                            );





                        });


                    it("Add Bcc recipients ",
                        function (done) {


                            /* ReadWriteItem or ReadWriteMailbox */
                            /* Add Bcc recipients (Applicable only on message) */
                            var newRecipients = [
                                {
                                    "displayName": "Paul Walker",
                                    "emailAddress": "paulw@contoso.com"
                                }
                            ];
                            Office.context.mailbox.item.bcc.addAsync(newRecipients,
                                function callback(asyncResult) {
                                    if (asyncResult.status == "failed") {
                                        console.log("Action failed with error: " + asyncResult.error.message);
                                        document.getElementById('addBccRecipients').innerHTML = ("Action failed with error: " + asyncResult.error.message);
                                    } else {
                                        console.log("Bcc add successfully");
                                        document.getElementById('addBccRecipients').innerHTML = ("Bcc Added successfully");

                                    }
                                    expect(asyncResult.status).toBe("succeeded");
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
                                    done();
                                }
                            );




                        });


                    it("Get Bcc recipients ",
                        function (done) {



                            /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                            /* Get Bcc recipients (Applicable only on message) */
                            Office.context.mailbox.item.bcc.getAsync(
                                function callback(asyncResult) {
                                    if (asyncResult.status == "failed") {
                                        console.log("Action failed with error: " + asyncResult.error.message);
                                        document.getElementById('getBccRecipients').innerHTML = ("Action failed with error: " + asyncResult.error.message);
                                    } else {
                                        var recipients = "";
                                        asyncResult.value.forEach(function (recipient, index) {
                                            recipients = recipients + recipient.displayName + " (" + recipient.emailAddress + ");";
                                        });
                                        console.log(recipients);
                                        document.getElementById('getBccRecipients').innerHTML = (recipients);
                                    }
                                    expect(asyncResult.status).toBe("succeeded");
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





                it("Get item type",
                    function (done) {



                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get item type */
                        console.log(Office.context.mailbox.item.itemType);
                        document.getElementById("itemType").innerHTML = Office.context.mailbox.item.itemType;
                        expect(Office.context.mailbox.item.itemType).toBeDefined();
                        done();
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




        });



    });

