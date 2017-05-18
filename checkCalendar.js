var alexa = require("alexa-app");
var app = new alexa.app("test");

var moment = require('moment');
var Moment = require('moment-timezone');

// Microsoft Graph JavaScript SDK
// npm install msgraph-sdk-javascript
var MicrosoftGraph = require("msgraph-sdk-javascript");

var delegateSlot = require("./index.js");
var response = require("./index.js");
var isSlot = require("./index.js");

//check Calendar
function checkCalendarIntent(request, session, callback){
    console.log("in calendar");
    console.log("request: "+JSON.stringify(request));
    var sessionAttributes={};
    var filledSlots = delegateSlot.delegateSlotCollection(request, sessionAttributes, callback);

    var startDate = request.intent.slots.startDate.value;
      // var endDate = request.intent.slots.endDate.value;

    var endDate = isSlot.isSlotValid(request, "endDate");
    if (endDate) {
      var url = '/me/calendar/calendarView?startDateTime='+ startDate +'T'+'00:00:00.0000000' + '&'+'endDateTime='+endDate+'T'+'23:59:59.0000000';
    } else {
      var url = '/me/calendar/calendarView?startDateTime='+ startDate +'T'+'00:00:00.0000000'  + '&'+'endDateTime='+startDate+'T'+'23:59:59.0000000';
    }

    //compose speechOutput that simply reads all the collected slot values
    var speechOutput='';

    console.log('session: '+JSON.stringify(session));
    var accessToken = session.user.accessToken;
    if(accessToken){
        // console.log('accessToken: ' + accessToken);
        var client = MicrosoftGraph.Client.init({
              authProvider: (done) => {
                  done(null, accessToken);
              }
        });

            // var Moment = require('moment-timezone');
            // var today = moment().format('YYYY-MM-DD');
            speakDate = startDate;

            console.log('type '+ typeof(startDate));
            console.log('startDate:' + startDate);
            console.log('endDate:' + endDate);
            //

        return client
            .api(url)
            .header("Prefer", 'outlook.timezone="Asia/Taipei"')
            .top(3)
            .select("subject")
            .select("start")
            .select("end")
            .get()
            .then((res) => {
              var upcomingEventNames = [];
              console.log(url);
              console.log(JSON.stringify( res));
              for (var i=0; i<res.value.length; i++) {
                  upcomingEventNames.push(JSON.stringify( res.value[i]));
              }

              var replyMessage = 'you have '+ upcomingEventNames.length +' meeting on ' + speakDate + ' ';

              for(var i=1; i<=upcomingEventNames.length; i++){
                  replyMessage += i+'. ' + res.value[i-1].subject + ' at ' + res.value[i-1].start.dateTime.substring(res.value[i-1].start.dateTime.lastIndexOf("T")+1,res.value[i-1].start.dateTime.lastIndexOf("."))+'. . ';
              }
              if(upcomingEventNames.length>=3){
                  replyMessage += 'for more, please check your calendar';
              }

              callback(sessionAttributes,
                  response.buildSpeechletResponse("mail status", replyMessage, "", false));
            }).catch((err) => {
                  console.log(err);
            });

    }else{
        console.log('no token');
    }

    //say the results
    // callback(sessionAttributes,
    //     buildSpeechletResponse("mail status", speechOutput, "", true));
}

exports.checkCalendarIntent = checkCalendarIntent;
