var alexa = require("alexa-app");
var app = new alexa.app("test");

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

    //compose speechOutput that simply reads all the collected slot values
    var speechOutput = "check mail now";

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

        // var url = '/me/mailFolders/';
          var url = '/me/messages';

          return   client
                  .api(url)
                  .header("Prefer", 'outlook.timezone="Asia/Taipei"')
                  .select("receivedDateTime")
                  .select("subject")
                  .select("bodyPreview")
                  .select("sender")
                  .top(3)
                  .get()
                  .then((res) => {

                    console.log(url);
                    console.log("check mail" + JSON.stringify(res));

                    var upcomingEvent = [];
                        // upcomingEventBodyPreview = [],
                        // upcomingEventReceivedDateTime =[];



                    // console.log('sender: ' + res.value[0].sender.emailAddress.name);

                    var temp = {};
                    var regex = new RegExp( mailSender, 'g' );
                    for(var i=0; i<res.value.length; i++) {
                      temp = {
                        subject: res.value[i].subject,
                        bodyPreview: res.value[i].bodyPreview,
                        receivedDateTime: res.value[i].receivedDateTime,
                        sender: res.value[i].sender.emailAddress.name
                      };
                      var str = res.value[i].sender.emailAddress.name;
                        if(str.match(regex)){
                          upcomingEvent.push(temp);
                          console.log("loop: " + res.value[i].sender.emailAddress.name);
                        }else {
                          console.log('err:' + res.value.length)
                        }
                      }

                      console.log("mail subject: " + upcomingEvent[0].subject);
                      console.log("mail bodyPreview: " + upcomingEvent[0].bodyPreview);
                      console.log("mail receivedDateTime: " + upcomingEvent[0].receivedDateTime);
                    for(var i = 0 ; i < upcomingEvent.length ; i++){
                      speechOutput += "mail: " + i + " subject " + upcomingEvent[i].subject + " content " + upcomingEvent[i].bodyPreview;
                    }


                    // speechOutput = "Receiver folder have unread mail " + upcomingEventNames.unreadItemCount + " and total mail " + upcomingEventNames.totalItemCount;
                    callback(sessionAttributes,
                        response.buildSpeechletResponse("mail status", speechOutput, "", true));

                  }).catch((err) =>{
                    console.log(err);
                  });

    }else{
        console.log('no token');
    }

    //say the results
    // callback(sessionAttributes,
    //     buildSpeechletResponse("mail status", speechOutput, "", true));
}

exports.checkMailIntent = checkMailIntent;
