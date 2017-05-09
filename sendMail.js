var alexa = require("alexa-app");
var app = new alexa.app("test");

// Microsoft Graph JavaScript SDK
// npm install msgraph-sdk-javascript
var MicrosoftGraph = require("msgraph-sdk-javascript");

var delegateSlot = require("./index.js");
var response = require("./index.js");

//send mail
function SendEmailIntent(request, session, callback){
    console.log("in mail");
    console.log("request: "+JSON.stringify(request));
    var sessionAttributes={};
    var filledSlots = delegateSlot.delegateSlotCollection(request, sessionAttributes, callback);

    //compose speechOutput that simply reads all the collected slot values
    var speechOutput = "Email is sent";

    //Now let's recap the trip
    var subject=request.intent.slots.mailSubject.value;
    var content=request.intent.slots.mailContent.value;
    var recipient = request.intent.slots.mailRecipient.value;

    speechOutput+= "send mail"  + " Recipient: " + recipient + " subject: " + subject + " content: "+ content;

    console.log('session: '+JSON.stringify(session));
    var accessToken = session.user.accessToken;
    if(subject && content && recipient){
      if(accessToken){
          // console.log('accessToken: ' + accessToken);
          var client = MicrosoftGraph.Client.init({
                authProvider: (done) => {
                    done(null, accessToken);
                }
          });

          // send mail
          const getmyContacts = () => new Promise((rs, rj) => {
              client.api('/me/contacts').get().then((contactsResult)=>{
                // get email address
                var eventContacts ={
                  name: '',
                  email: ''
                }

                  for (var i=0; i<contactsResult.value.length; i++) {
                      if(contactsResult.value[i].givenName == recipient){
                        eventContacts.name = contactsResult.value[i].givenName;
                        eventContacts.email = contactsResult.value[i].emailAddresses[0].address;
                        console.log("compare: " + contactsResult.value[i].givenName);
                      }else{
                        console.log("no pair");
                        console.log('res name: ' + contactsResult.value[i].givenName);
                        console.log('res address: ' + contactsResult.value[i].emailAddresses[0].address);
                      }
                    }

                // if(eventContacts.name && eventContacts.email){
                //   speechOutput = " mail name: " + eventContacts.name + " mail address: " + eventContacts.email;
                // }else{
                //   speechOutput = 'no user to check'
                // }

                rs(eventContacts);
              }).catch((e) => {
                rj(e);
              })
            });


          const sendEmail = (eventContacts) => new Promise((rs, rj) => {
            // handle contactsResult
            var mailAddress = eventContacts.email;
            var mail = {
                subject: subject,
                toRecipients: [{
                    emailAddress: {
                        address: mailAddress
                    }
                }],
                body: {
                    content: content,
                    contentType: "html"
                }
            }
            client.api('/me/sendMail').post({message:mail}).then((mailResult)=>{
              console.log(mail);
              rs(mailResult);
            }).catch((e) => {
              rj(e);
            })
          });

          getmyContacts()
            .then(sendEmail)
            .then((mailResult)=>{
              // do something
              callback(sessionAttributes,
                  response.buildSpeechletResponse("mail status", speechOutput, "", true));
            }).catch((e) => {
              console.error('error happen', e);
            })

        // return client
        //     .api('/me/sendMail')
        //     .post({message:mail})
        //     .then((res) => {
        //       // console.log('request content' + JSON.stringify(request) );
        //       // console.log('res content' + JSON.stringify(res) );
        //       // console.log('response content' + JSON.stringify(response) );
        //       // response.say("send an mail title: "+ title +' now content: ' + content).reprompt("please say again").shouldEndSession(false);
        //       // templateSubject = '';
        //       // templateContent = ''
        //       callback(sessionAttributes,
        //           buildSpeechletResponse("mail status", speechOutput, "", true));
        //     }).catch((err) => {
        //       console.log(err);
        //     });

      }else{
          console.log('no token');
      }
    }else{
      console.log("no subject, no content");
    }
    //say the results
    // callback(sessionAttributes,
    //     buildSpeechletResponse("mail status", speechOutput, "", true));
}

exports.SendEmailIntent = SendEmailIntent;
