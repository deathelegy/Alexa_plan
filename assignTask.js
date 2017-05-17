var delegateSlot = require("./index.js");
var response = require("./index.js");
var buildSpeechletResponseWithDirectiveNoIntent = require("./index.js");

//AssignTaskIntent
function AssignTaskIntent(request, session, callback){
    console.log("in assign task");
    console.log("request: "+JSON.stringify(request));

    console.log("request been delete :  "+JSON.stringify(request));
    var sessionAttributes={};
    var filledSlots = slotCollection(request, sessionAttributes, callback);

    //compose speechOutput that simply reads all the collected slot values
    var speechOutput = "AssignTask now";

    //Now let's recap the trip slot
    var recipient=request.intent.slots.recipient.value;
    var thing=request.intent.slots.thing.value;
    var duedate=request.intent.slots.duedate.value;

      speechOutput+= " information:"  + " recipient: "+ recipient + " thing: "+ thing + " duedate: " + duedate;
      var replyMessage = '.. Is there anything else I can help you with?'
      speechOutput += replyMessage;
      console.log('session: '+JSON.stringify(session));

      //say the results
      callback(sessionAttributes, response.buildSpeechletResponse("AssignTask status", speechOutput, "", false));
}

//delegateSlot

function slotCollection(request, sessionAttributes, callback){
  console.log("in delegateSlotCollection");
  console.log("  current dialogState: "+JSON.stringify(request.dialogState));

    if (request.dialogState === "STARTED") {
      console.log("AssignTask in started");
      console.log("  current request: "+JSON.stringify(request));
      console.log("  in started dialogState: "+JSON.stringify(request.dialogState));
      var updatedIntent=request.intent;
      //optionally pre-fill slots: update the intent object with slot values for which
      //you have defaults, then return Dialog.Delegate with this updated intent
      // in the updatedIntent property
      callback(sessionAttributes,
          buildSpeechletResponseWithDirectiveNoIntent.buildSpeechletResponseWithDirectiveNoIntent());
    } else if (request.dialogState !== "COMPLETED") {
      console.log("AssignTask in not completed");
      console.log("  current request: "+JSON.stringify(request));
      console.log("in not completed dialogState: "+JSON.stringify(request.dialogState));

      if(request.intent.slots.recipient.value == 'Kaya'){
        // console.log("recipient:" + request.intent.slots.recipient.value);
        // delete request.intent.slots.recipient.value;
        // console.log("recipient has been delete :" + request.intent.slots.recipient.value);
        var speechOutput = "AssignTask now : OH NO";
        callback(sessionAttributes, response.buildSpeechletResponse("AssignTask status", speechOutput, "", false));
      }

      // return a Dialog.Delegate directive with no updatedIntent property.
      callback(sessionAttributes,
          buildSpeechletResponseWithDirectiveNoIntent.buildSpeechletResponseWithDirectiveNoIntent());
    } else {
      console.log("AssignTask  in completed");
      console.log("  current request: "+JSON.stringify(request));
      console.log("  returning: "+ JSON.stringify(request.intent));
      console.log("in completed dialogState: "+ JSON.stringify(request.dialogState));
      // Dialog is now complete and all required slots should be filled,
      // so call your normal intent handler.
      return request.intent;
    }
}

exports.AssignTaskIntent = AssignTaskIntent;
