var alexa = require("alexa-app");
var app = new alexa.app("test");
var moment = require('moment');
var Moment = require('moment-timezone');

// Microsoft Graph JavaScript SDK
// npm install msgraph-sdk-javascript
var MicrosoftGraph = require("msgraph-sdk-javascript");

var delegateSlot = require("./index.js");
var response = require("./index.js");

//check mail
function checkMailIntent(request, session, callback){
    console.log("in mail box");
    console.log("request: "+JSON.stringify(request));
    var sessionAttributes={};
    var filledSlots = delegateSlot.delegateSlotCollection(request, sessionAttributes, callback);

    //get today
    var today = moment().format();
    console.log("today:  "+ today);
    //compose speechOutput that simply reads all the collected slot values
    var speechOutput='';

    var mailSender = request.intent.slots.mailSender.value;

    console.log('session: '+JSON.stringify(session));
    var accessToken = session.user.accessToken;
    if(accessToken){
        // console.log('accessToken: ' + accessToken);
        var client = MicrosoftGraph.Client.init({
              authProvider: (done) => {
                  done(null, accessToken);
              }
        });

        //get folderID

        const getfolderID = () => new Promise((rs, rj) => {
            client.api('/me/mailFolders/').get().then((folderIDResult)=>{
              var folderID = '';

              for (var i=0; i<folderIDResult.value.length; i++) {
                  if(folderIDResult.value[i].displayName == '收件匣'){
                    folderID = folderIDResult.value[i].id;
                    console.log("compare: " + folderID);
                  }else{
                    console.log("no pair folderID");
                  }
                }

              rs(folderID);
            }).catch((e) => {
              rj(e);
            })
          });

       // list folder messages

       const listFolderMessage = (folderID) => new Promise((rs, rj) => {
         // handle contactsResult
         client
                .api('/me/mailFolders/'+ folderID +'/messages/')
                .select("receivedDateTime")
                .select("subject")
                .select("bodyPreview")
                .select("sender")
                .top(10)
                .get().then((folderMessageResult)=>{

                  console.log("check mail" + JSON.stringify(folderMessageResult));

                  var upcomingEvent = [];

                  // console.log('sender: ' + res.value[0].sender.emailAddress.name);
                  mailSender = mailSender.toLowerCase()
                  var temp = {};
                  var regex = new RegExp( mailSender, 'g' );
                  for(var i=0; i<folderMessageResult.value.length; i++) {
                    var tempTaipeiTime = moment(folderMessageResult.value[i].receivedDateTime).add(8,"hours").format();
                    temp = {
                      subject: folderMessageResult.value[i].subject,
                      bodyPreview: folderMessageResult.value[i].bodyPreview,
                      receivedDateTime: tempTaipeiTime,
                      sender: folderMessageResult.value[i].sender.emailAddress.name
                    };

                    console.log("time test: "+ temp.receivedDateTime);
                    var str = folderMessageResult.value[i].sender.emailAddress.name.toLowerCase();
                      if(str.match(regex) && moment(temp.receivedDateTime).isSame(today ,'day')){
                        upcomingEvent.push(temp);
                        console.log("loop: " + folderMessageResult.value[i].sender.emailAddress.name);
                        console.log("loop: " +　JSON.stringify(upcomingEvent));
                      }else {
                        console.log('err:' + folderMessageResult.value.length)
                      }
                    }

                  if(upcomingEvent.length > 0){
                    speechOutput += "You have " + upcomingEvent.length + " .. mail from " + mailSender ;
                    for(var i = 1 ; i <= upcomingEvent.length ; i++){
                      speechOutput += ".. mail: " + i + " .. subject " + upcomingEvent[i-1].subject + " .. content " + upcomingEvent[i-1].bodyPreview;
                    }
                  }else{
                    speechOutput += "sorry, you don't have any mail today from " + mailSender;
                  }

           rs(folderMessageResult);
         }).catch((e) => {
           rj(e);
         })
       });

       // do

       getfolderID()
         .then(listFolderMessage)
         .then((folderMessageResult)=>{

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

    //say the results
    // callback(sessionAttributes,
    //     buildSpeechletResponse("mail status", speechOutput, "", true));
}

exports.checkMailIntent = checkMailIntent;
