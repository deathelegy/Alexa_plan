var alexa = require("alexa-app");
var app = new alexa.app("test");

// Microsoft Graph JavaScript SDK
// npm install msgraph-sdk-javascript
var MicrosoftGraph = require("msgraph-sdk-javascript");

var delegateSlot = require("./index.js");
var response = require("./index.js");
var count = 0;
//send mail
function SendEmailIntent(request, session, callback){
    console.log("in mail");
    console.log("request: "+JSON.stringify(request));
    var sessionAttributes={};
    var filledSlots = delegateSlot.delegateSlotCollection(request, sessionAttributes, callback);

    //compose speechOutput that simply reads all the collected slot values
    var speechOutput = "";

    //Now let's recap the trip
    var subject=request.intent.slots.mailSubject.value;
    var content=request.intent.slots.mailContent.value;
    var recipient = request.intent.slots.mailRecipient.value;

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

                console.log(JSON.stringify(contactsResult));
                // get email address

                var eventContacts = [];

                var temp={};
                recipient = recipient.toLowerCase()
                var regex = new RegExp( recipient, 'g' );

                  for (var i=0; i<contactsResult.value.length; i++) {
                    var str = contactsResult.value[i].givenName.toLowerCase();
                      if(str.match(regex)){
                        temp = {
                            name: contactsResult.value[i].givenName,
                            email: contactsResult.value[i].emailAddresses[0].address,
                            last_name: contactsResult.value[i].surname
                          };
                        eventContacts.push(temp);
                        console.log("compare: " + contactsResult.value[i].givenName);
                        console.log("compare mail : " + contactsResult.value[i].emailAddresses[0].address);
                      }else{
                        console.log("no pair");
                        console.log('res name: ' + contactsResult.value[i].givenName);
                        console.log('res address: ' + contactsResult.value[i].emailAddresses[0].address);

                      }
                    }

                    if(eventContacts.length < 1){
                      speechOutput +='sorry can not find '+ recipient + 'in your contacts ..please try anoter name';

                      callback(sessionAttributes,
                          response.buildSpeechletResponse("mail status", speechOutput, "", false));

                    }else if(eventContacts.length > 1){
                      speechOutput +='sorry you have '+ eventContacts.length + ' same names .. ' + recipient +' in your contacts ..please select correct one';
                      for(var i = 1 ; i <= eventContacts.length ; i++){
                        speechOutput +='. . contacts' + i +' . . '+ eventContacts[i-1].name +' . . ' +eventContacts[i-1].last_name;
                      }

                      callback(sessionAttributes,
                          response.buildSpeechletResponse("mail status", speechOutput, "", false));
                      // request.intent.slots.mailRecipient.value = '';
                      // filledSlots = delegateSlot.delegateSlotCollection(request, sessionAttributes, callback);
                      // recipient = request.intent.slots.mailRecipient.value;
                      // rj(e);
                    }else {
                      rs(eventContacts);
                    }                
              }).catch((e) => {
                rj(e);
              })
            });


          const sendEmail = (eventContacts) => new Promise((rs, rj) => {
            // handle contactsResult
            count++;
            console.log('time: ' + count);
            var mailAddress = eventContacts[0].email;
            console.log(mailAddress);
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

            client
            .api('/me/sendMail')
            .post({message:mail})
            .then((mailResult)=>{
              console.log(JSON.stringify(mail));
              rs(mailResult);
            }).catch((e) => {
              rj(e);
            })
          });

          getmyContacts()
            .then(sendEmail)
            .then((mailResult)=>{

              speechOutput+= "send mail"  + " Recipient: " + recipient + " subject: " + subject + " content: "+ content ;
              var replyMessage = '.. Is there anything else I can help you with?'
              speechOutput += replyMessage;
              // do something
              callback(sessionAttributes,
                  response.buildSpeechletResponse("mail status", speechOutput, "", false));
            }).catch((e) => {
              console.error('error happen', e);
            })

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
