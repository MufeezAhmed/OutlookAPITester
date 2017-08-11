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


      


       


  

      
        describe("1.5 API Office.context.", function () {


                it(" close Container :Commented to validate rest of the test cases",
                    function () {

                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* close Container */
                         Office.context.ui.closeContainer();
                       
                      
                        // document.getElementById("inlineImageDisplayReplyForm").innerHTML = "Use Read Test Addin ";
                        //document.getElementById("inlineImageDisplayReplyAllForm").innerHTML = "Use Read Test Addin ";
                    });
            

            


               




            });
       
      });
       
  

