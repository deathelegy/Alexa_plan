var alexa = require("alexa-app");
var app = new alexa.app("test");

// Microsoft Graph JavaScript SDK
// npm install msgraph-sdk-javascript
var MicrosoftGraph = require("msgraph-sdk-javascript");

var assignTask = require("./assignTask.js");
var sendMail = require("./sendMail.js");
var checkMail = require("./checkMail.js");
var checkCalendar = require("./checkCalendar.js");

'use strict';

/**
 * This sample demonstrates a simple skill built with the Amazon Alexa Skills Kit.
 * The Intent Schema, Custom Slots, and Sample Utterances for this skill, as well as
 * testing instructions are located at http://amzn.to/1LzFrj6
 *
 * For additional samples, visit the Alexa Skills Kit Getting Started guide at
 * http://amzn.to/1LGWsLG
 */
 var speechOutput;
 var reprompt;
 var welcomeOutput = "welcome mail service, what do you want to do?";
 var welcomeReprompt = "Let me know what do you want to do?";
 var tripIntro = [
   "This sounds like a cool trip. ",
   "This will be fun. ",
   "Oh, I like this trip. "
 ];

// --------------- Helpers that build all of the responses -----------------------

function buildSpeechletResponse(title, output, repromptText, shouldEndSession) {
    return {
        outputSpeech: {
            type: 'PlainText',
            text: output,
        },
        card: {
            type: 'Simple',
            title: `SessionSpeechlet - ${title}`,
            content: `SessionSpeechlet - ${output}`,
        },
        reprompt: {
            outputSpeech: {
                type: 'PlainText',
                text: repromptText,
            },
        },
        shouldEndSession,
    };
}

function buildResponse(sessionAttributes, speechletResponse) {
    console.log("Responding with " + JSON.stringify(speechletResponse));
    return {
        version: '1.0',
        sessionAttributes,
        response: speechletResponse,
    };
}

function buildSpeechletResponseWithDirectiveNoIntent() {
    console.log("in buildSpeechletResponseWithDirectiveNoIntent");
    return {
      "outputSpeech" : null,
      "card" : null,
      "directives" : [ {
        "type" : "Dialog.Delegate"
      } ],
      "reprompt" : null,
      "shouldEndSession" : false
    }
  }

  function buildSpeechletResponseDelegate(shouldEndSession) {
      return {
          outputSpeech:null,
          directives: [
                  {
                      "type": "Dialog.Delegate",
                      "updatedIntent": null
                  }
              ],
         reprompt: null,
          shouldEndSession: shouldEndSession
      }
  }


// --------------- Functions that control the skill's behavior -----------------------

function getWelcomeResponse(callback) {
    console.log("in welcomeResponse");
    // If we wanted to initialize the session to have some attributes we could add those here.
    const sessionAttributes = {};
    const cardTitle = 'Welcome';
    const speechOutput = 'welcome mail service, what do you want to do?';
    // If the user either does not reply to the welcome message or says something that is not
    // understood, they will be prompted again with this text.
    const repromptText = "welcome mail service, what do you want to do?";
    const shouldEndSession = false;

    callback(sessionAttributes,
        buildSpeechletResponse(cardTitle, speechOutput, repromptText, shouldEndSession));

}

//get contacts
// function getContacts(request, session, callback){
//     console.log("in contacts");
//     console.log("request: "+JSON.stringify(request));
//     var sessionAttributes={};
//     var filledSlots = delegateSlotCollection(request, sessionAttributes, callback);
//
//     //compose speechOutput that simply reads all the collected slot values
//     var speechOutput = "list contacts";
//     var mailName = request.intent.slots.mailName.value;
//
//     console.log('session: '+JSON.stringify(session));
//     var accessToken = session.user.accessToken;
//
//       if(accessToken && mailName){
//           // console.log('accessToken: ' + accessToken);
//           var client = MicrosoftGraph.Client.init({
//                 authProvider: (done) => {
//                     done(null, accessToken);
//                 }
//           });
//           //
//           //to who
//           var url = '/me/contacts';
//
//
//           return client
//               .api('/me/contacts')
//               .get()
//               .then((res) => {
//                 // console.log('request content' + JSON.stringify(request) );
//                 // console.log('res content' + JSON.stringify(res) );
//                 // console.log('res name: ' + res.value[0].givenName);
//                 // console.log('res address: ' + res.value[0].emailAddresses[0].address);
//                 // console.log('res length: ' + res.value.length);
//                 // console.log('mailName:' + mailName);
//                 // console.log('mailName type:' + typeof(mailName));
//                 // console.log('res name type: ' + typeof(res.value[0].givenName));
//
//                 var eventContacts ={
//                   name: '',
//                   email: ''
//                 }
//
//                   for (var i=0; i<res.value.length; i++) {
//                       if(res.value[i].givenName == mailName){
//                         eventContacts.name = res.value[i].givenName;
//                         eventContacts.email = res.value[i].emailAddresses[0].address;
//                         console.log("compare: " + res.value[i].givenName);
//                       }else{
//                         console.log("no pair");
//                         console.log('res name: ' + res.value[i].givenName);
//                         console.log('res address: ' + res.value[i].emailAddresses[0].address);
//                       }
//                     }
//
//                 if(eventContacts.name && eventContacts.email){
//                   speechOutput = " mail name: " + eventContacts.name + " mail address: " + eventContacts.email;
//                 }else{
//                   speechOutput = 'no user to check'
//                 }
//
//                 callback(sessionAttributes,
//                     buildSpeechletResponse("contacts status", speechOutput, "", true));
//
//               }).catch((err) => {
//                 console.log(err);
//               });
//
//       }else{
//           console.log('no token');
//       }
//
//     //say the results
//     // callback(sessionAttributes,
//     //     buildSpeechletResponse("mail status", speechOutput, "", true));
// }


function handleSessionEndRequest(callback) {
    const cardTitle = 'Session Ended';
    const speechOutput = 'Thank you for using this mail service. Have a nice day!';
    // Setting this to true ends the session and exits the skill.
    const shouldEndSession = true;

    callback({}, buildSpeechletResponse(cardTitle, speechOutput, null, shouldEndSession));
}

function delegateSlotCollection(request, sessionAttributes, callback){
  console.log("in delegateSlotCollection");
  console.log("  current dialogState: "+JSON.stringify(request.dialogState));

    if (request.dialogState === "STARTED") {
      console.log("in started");
      console.log("  current request: "+JSON.stringify(request));
      var updatedIntent=request.intent;
      //optionally pre-fill slots: update the intent object with slot values for which
      //you have defaults, then return Dialog.Delegate with this updated intent
      // in the updatedIntent property
      callback(sessionAttributes,
          buildSpeechletResponseWithDirectiveNoIntent());
    } else if (request.dialogState !== "COMPLETED") {
      console.log("in not completed");
      console.log("  current request: "+JSON.stringify(request));
      // return a Dialog.Delegate directive with no updatedIntent property.
      callback(sessionAttributes,
          buildSpeechletResponseWithDirectiveNoIntent());
    } else {
      console.log("in completed");
      console.log("  current request: "+JSON.stringify(request));
      console.log("  returning: "+ JSON.stringify(request.intent));
      // Dialog is now complete and all required slots should be filled,
      // so call your normal intent handler.
      return request.intent;
    }
}

function randomPhrase(array) {
    // the argument is an array [] of words or phrases
    var i = 0;
    i = Math.floor(Math.random() * array.length);
    return(array[i]);
}
function isSlotValid(request, slotName){
        var slot = request.intent.slots[slotName];
        //console.log("request = "+JSON.stringify(request)); //uncomment if you want to see the request
        var slotValue;

        //if we have a slot, get the text and store it into speechOutput
        if (slot && slot.value) {
            //we have a value in the slot
            slotValue = slot.value.toLowerCase();
            return slotValue;
        } else {
            //we didn't get a value in the slot.
            return false;
        }
}


// --------------- Events -----------------------

/**
 * Called when the session starts.
 */
function onSessionStarted(sessionStartedRequest, session) {
    console.log(`onSessionStarted requestId=${sessionStartedRequest.requestId}, sessionId=${session.sessionId}`);
}

/**
 * Called when the user launches the skill without specifying what they want.
 */
function onLaunch(request, session, callback) {
    //console.log(`onLaunch requestId=${launchRequest.requestId}, sessionId=${session.sessionId}`);
    console.log("in launchRequest");
    console.log("  request: "+JSON.stringify(request));
    // Dispatch to your skill's launch.
    getWelcomeResponse(callback);
}

/**
 * Called when the user specifies an intent for this skill.
 */
function onIntent(request, session, callback) {
    //console.log(`onIntent requestId=${intentRequest.requestId}, sessionId=${session.sessionId}`);
    console.log("in onIntent");
    console.log("  request: "+JSON.stringify(request));

    const intent = request.intent;
    const intentName = request.intent.name;

    // Dispatch to your skill's intent handlers
    if (intentName === 'SendEmailIntent') {
        sendMail.SendEmailIntent(request, session, callback);
    }else if(intentName === 'checkMailIntent'){
      checkMail.checkMailIntent(request, session, callback);
    }else if(intentName === 'AssignTaskIntent'){
      assignTask.AssignTaskIntent(request, session, callback);
    }else if(intentName === 'checkCalendarIntent'){
      checkCalendar.checkCalendarIntent(request, session, callback);
    }else if(intentName === 'getContacts'){
      getContacts(request, session, callback);
    }else if (intentName === 'AMAZON.HelpIntent') {
        getWelcomeResponse(callback);
    } else if (intentName === 'AMAZON.StopIntent' || intentName === 'AMAZON.CancelIntent') {
        handleSessionEndRequest(callback);
    } else {
        throw new Error('Invalid intent');
    }
}

/**
 * Called when the user ends the session.
 * Is not called when the skill returns shouldEndSession=true.
 */
function onSessionEnded(sessionEndedRequest, session) {
    console.log(`onSessionEnded requestId=${sessionEndedRequest.requestId}, sessionId=${session.sessionId}`);
    // Add cleanup logic here
}


// --------------- Main handler -----------------------

// Route the incoming request based on type (LaunchRequest, IntentRequest,
// etc.) The JSON body of the request is provided in the event parameter.
exports.handler = (event, context, callback) => {
    try {
        // console.log(`event.session.application.applicationId=${event.session.application.applicationId}`);
        console.log("EVENT=====" + JSON.stringify(event));

        /**
         * Uncomment this if statement and populate with your skill's application ID to
         * prevent someone else from configuring a skill that sends requests to this function.
         */
        /*
        if (event.session.application.applicationId !== 'amzn1.echo-sdk-ams.app.[unique-value-here]') {
             callback('Invalid Application ID');
        }
        */

        if (event.session.new) {
            onSessionStarted({ requestId: event.request.requestId }, event.session);
        }

        if (event.request.type === 'LaunchRequest') {
            onLaunch(event.request,
                event.session,
                (sessionAttributes, speechletResponse) => {
                    callback(null, buildResponse(sessionAttributes, speechletResponse));
                });
        } else if (event.request.type === 'IntentRequest') {
            onIntent(event.request,
                event.session,
                (sessionAttributes, speechletResponse) => {
                    callback(null, buildResponse(sessionAttributes, speechletResponse));
                });
        } else if (event.request.type === 'SessionEndedRequest') {
            onSessionEnded(event.request, event.session);
            callback();
        }
    } catch (err) {
        callback(err);
    }
};

exports.isSlotValid = isSlotValid;
exports.delegateSlotCollection = delegateSlotCollection;
exports.buildSpeechletResponse = buildSpeechletResponse;
