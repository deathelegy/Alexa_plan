// var alexa = require("alexa-app");
// var app = new alexa.app("test");

// Microsoft Graph JavaScript SDK
// npm install msgraph-sdk-javascript
var MicrosoftGraph = require("msgraph-sdk-javascript");

var delegateSlot = require("./index.js");
var response = require("./index.js");
var isSlot = require("./index.js");
var buildSpeechletResponseWithDirectiveNoIntent = require("./index.js");

var count = 0;
var eventContacts = [];
var speechOutput = '';


const compare = (recipient, lastName, contactsResult) => new Promise((rs, rj) => {

  eventContacts = [];
  var temp={};
  recipient = recipient.toLowerCase()
  var regexFirst = new RegExp( recipient, 'g' );
  var regexLast = new RegExp( lastName, 'g' );

  if(lastName){
    console.log('in lastName status:')
    for (var i=0; i<contactsResult.value.length; i++) {
      var str_first = contactsResult.value[i].givenName.toLowerCase();
      var str_last = contactsResult.value[i].surname.toLowerCase();
        if(str_first.match(regexFirst) && str_last.match(regexLast)){
          temp = {
              name: contactsResult.value[i].givenName,
              email: contactsResult.value[i].emailAddresses[0].address,
              last_name: contactsResult.value[i].surname
            };
          eventContacts.push(temp);
          console.log("compare first: " + contactsResult.value[i].givenName);
          console.log("compare last: " + contactsResult.value[i].surname);
          console.log("compare mail : " + contactsResult.value[i].emailAddresses[0].address);
        }else{
          console.log("no pair");
          console.log('res name: ' + contactsResult.value[i].givenName);
          console.log('res last name: ' + contactsResult.value[i].surname);
          console.log('res address: ' + contactsResult.value[i].emailAddresses[0].address);
        }
      }
  }else{
    console.log('in firstName status:')
    for (var i=0; i<contactsResult.value.length; i++) {
      var str = contactsResult.value[i].givenName.toLowerCase();
        if(str.match(regexFirst)){
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
    }

  });

// compare
// function compare(recipient, lastName, contactsResult){
//   eventContacts = [];
//   var temp={};
//   recipient = recipient.toLowerCase()
//   var regexFirst = new RegExp( recipient, 'g' );
//   var regexLast = new RegExp( lastName, 'g' );
//
//   if(lastName){
//     console.log('in lastName status:')
//     for (var i=0; i<contactsResult.value.length; i++) {
//       var str_first = contactsResult.value[i].givenName.toLowerCase();
//       var str_last = contactsResult.value[i].surname.toLowerCase();
//         if(str_first.match(regexFirst) && str_last.match(regexLast)){
//           temp = {
//               name: contactsResult.value[i].givenName,
//               email: contactsResult.value[i].emailAddresses[0].address,
//               last_name: contactsResult.value[i].surname
//             };
//           eventContacts.push(temp);
//           console.log("compare first: " + contactsResult.value[i].givenName);
//           console.log("compare last: " + contactsResult.value[i].surname);
//           console.log("compare mail : " + contactsResult.value[i].emailAddresses[0].address);
//         }else{
//           console.log("no pair");
//           console.log('res name: ' + contactsResult.value[i].givenName);
//           console.log('res last name: ' + contactsResult.value[i].surname);
//           console.log('res address: ' + contactsResult.value[i].emailAddresses[0].address);
//         }
//       }
//       return eventContacts;
//   }else{
//     console.log('in firstName status:')
//     for (var i=0; i<contactsResult.value.length; i++) {
//       var str = contactsResult.value[i].givenName.toLowerCase();
//         if(str.match(regexFirst)){
//           temp = {
//               name: contactsResult.value[i].givenName,
//               email: contactsResult.value[i].emailAddresses[0].address,
//               last_name: contactsResult.value[i].surname
//             };
//           eventContacts.push(temp);
//           console.log("compare: " + contactsResult.value[i].givenName);
//           console.log("compare mail : " + contactsResult.value[i].emailAddresses[0].address);
//         }else{
//           console.log("no pair");
//           console.log('res name: ' + contactsResult.value[i].givenName);
//           console.log('res address: ' + contactsResult.value[i].emailAddresses[0].address);
//         }
//       }
//     return eventContacts;
//   }
// }

//delegateSlot

function slotCollection(request, session, sessionAttributes, callback){
  console.log("in SlotCollection");
  console.log("  current dialogState: "+JSON.stringify(request.dialogState));
  console.log('session: '+JSON.stringify(session));
  var accessToken = session.user.accessToken;

    if (request.dialogState === "STARTED") {
      console.log("mail in started");
      console.log("  current request: "+JSON.stringify(request));
      console.log("  in started dialogState: "+JSON.stringify(request.dialogState));
      var updatedIntent=request.intent;
      //optionally pre-fill slots: update the intent object with slot values for which
      //you have defaults, then return Dialog.Delegate with this updated intent
      // in the updatedIntent property
      callback(sessionAttributes,
          buildSpeechletResponseWithDirectiveNoIntent.buildSpeechletResponseWithDirectiveNoIntent());
    } else if (request.dialogState !== "COMPLETED") {
      console.log("mail in not completed");
      console.log("  current request: "+JSON.stringify(request));
      console.log("in not completed dialogState: "+JSON.stringify(request.dialogState));
      var tempRecipient = request.intent.slots.mailRecipient.value;
      var tempLastName = request.intent.slots.mailRecipientLastName.value;

      if(tempRecipient){
        if(accessToken){
          var client = MicrosoftGraph.Client.init({
                authProvider: (done) => {
                    done(null, accessToken);
                }
          });

          const getmyContacts = () => new Promise((rs, rj) => {
              client.api('/me/contacts').get().then((contactsResult)=>{

                console.log("getmyContacts: " + JSON.stringify(contactsResult));
                // get email address
                compare(tempRecipient, tempLastName, contactsResult);
                rs(contactsResult);
              }).catch((e) => {
                rj(e);
              })
            });

            getmyContacts().then((eventContacts) => {
              console.log("out: " + JSON.stringify(eventContacts));
              if(eventContacts.length < 1){
                speechOutput +='sorry can not find '+ tempRecipient + ' in your contacts ..please try anoter name';

                callback(sessionAttributes,
                    response.buildSpeechletResponse("mail status", speechOutput, "", false));
              }else if(eventContacts.length > 1){
                speechOutput +='sorry you have '+ eventContacts.length + ' same names .. ' + tempRecipient +' in your contacts ..please select correct one';
                for(var i = 1 ; i <= eventContacts.length ; i++){
                  speechOutput +='. . contacts' + i +' . . '+ eventContacts[i-1].name +' . . ' +eventContacts[i-1].last_name;
                }
                callback(sessionAttributes,
                    response.buildSpeechletResponse("mail status", speechOutput, "", false));
              }else {
                console.log("good compare" + JSON.stringify(eventContacts));
              }
              eventContacts = '';
            }).catch((e) => {
              console.error('error happen', e);
            })
      }
    }
      // return a Dialog.Delegate directive with no updatedIntent property.
      callback(sessionAttributes,
          buildSpeechletResponseWithDirectiveNoIntent.buildSpeechletResponseWithDirectiveNoIntent());
    } else {
      console.log("mail  in completed");
      console.log("  current request: "+JSON.stringify(request));
      console.log("  returning: "+ JSON.stringify(request.intent));
      console.log("in completed dialogState: "+ JSON.stringify(request.dialogState));
      // Dialog is now complete and all required slots should be filled,
      // so call your normal intent handler.
      return request.intent;
    }
}

//send mail intent
function SendEmailIntent(request, session, callback){
    console.log("in mail");
    console.log("request: "+JSON.stringify(request));
    var sessionAttributes={};
    var filledSlots = slotCollection(request, session, sessionAttributes, callback);

    //compose speechOutput that simply reads all the collected slot values
    var speechOutput = "";

    //Now let's recap the trip
    var subject=request.intent.slots.mailSubject.value;
    var content=request.intent.slots.mailContent.value;

    if(eventContacts.length >0){
      if(subject && content){
        console.log("subject: " + subject);
        console.log("content: " + content);
        console.log("in send mail intent: " + JSON.stringify(eventContacts));

        callback(sessionAttributes,
            response.buildSpeechletResponse("mail status", "good to send", "", false));
      }
    }



    // var recipient = request.intent.slots.mailRecipient.value;
    //
    // var lastName = isSlot.isSlotValid(request, "mailRecipientLastName");
    //
    // console.log("LastName: " + lastName);
    //
    // console.log('session: '+JSON.stringify(session));
    // var accessToken = session.user.accessToken;
    // if(subject && content && recipient){
    //   if(accessToken){
    //       // console.log('accessToken: ' + accessToken);
    //       var client = MicrosoftGraph.Client.init({
    //             authProvider: (done) => {
    //                 done(null, accessToken);
    //             }
    //       });
    //
    //       //get contacts
    //
    //       const getmyContacts = () => new Promise((rs, rj) => {
    //           client.api('/me/contacts').get().then((contactsResult)=>{
    //
    //             console.log(JSON.stringify(contactsResult));
    //             // get email address
    //
    //             compare(recipient, lastName, contactsResult);
    //
    //                 if(eventContacts.length < 1){
    //                   speechOutput +='sorry can not find '+ recipient + ' in your contacts ..please try anoter name';
    //
    //                   callback(sessionAttributes,
    //                       response.buildSpeechletResponse("mail status", speechOutput, "", false));
    //
    //                 }else if(eventContacts.length > 1){
    //                   speechOutput +='sorry you have '+ eventContacts.length + ' same names .. ' + recipient +' in your contacts ..please select correct one';
    //                   for(var i = 1 ; i <= eventContacts.length ; i++){
    //                     speechOutput +='. . contacts' + i +' . . '+ eventContacts[i-1].name +' . . ' +eventContacts[i-1].last_name;
    //                   }
    //
    //                   callback(sessionAttributes,
    //                       response.buildSpeechletResponse("mail status", speechOutput, "", false));
    //
    //                   // console.log("1 recipient :" + request.intent.slots.mailRecipient.value);
    //                   // // request.intent.slots.mailRecipient.value = '';
    //                   // // request.intent.slots.mailRecipientLastName.value = '';
    //                   // delete request.intent.slots.mailRecipient.value;
    //                   // delete request.intent.slots.mailRecipientLastName.value;
    //                   // request.dialogState = 'IN_PROGRESS';
    //                   // FilledSlots = delegateSlot.delegateSlotCollection(request, sessionAttributes, callback);
    //                   //
    //                   //
    //                   // recipient  = request.intent.slots.mailRecipient.value;
    //                   // lastName = request.intent.slots.mailRecipientLastName.value;
    //                   // console.log("2 recipient :" + recipient);
    //                   // console.log(" request: "+JSON.stringify(request));
    //                 }else {
    //                   rs(eventContacts);
    //                 }
    //           }).catch((e) => {
    //             rj(e);
    //           })
    //         });
    //

            // send mail
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
    //
    //
    //       getmyContacts()
    //         .then(sendEmail)
    //         .then((mailResult)=>{
    //
    //           speechOutput+= "send mail"  + " Recipient: " + recipient + " subject: " + subject + " content: "+ content ;
    //           var replyMessage = '.. Is there anything else I can help you with?'
    //           speechOutput += replyMessage;
    //           // do something
    //           callback(sessionAttributes,
    //               response.buildSpeechletResponse("mail status", speechOutput, "", false));
    //         }).catch((e) => {
    //           console.error('error happen', e);
    //         })
    //
    //   }else{
    //       console.log('no token');
    //   }
    // }else{
    //   console.log("no subject, no content");
    // }
    //say the results
    // callback(sessionAttributes,
    //     buildSpeechletResponse("mail status", speechOutput, "", true));
}

exports.SendEmailIntent = SendEmailIntent;
