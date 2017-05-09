var delegateSlot = require("./index.js");
var response = require("./index.js");

//AssignTaskIntent
function AssignTaskIntent(request, session, callback){
    console.log("in assign task");
    console.log("request: "+JSON.stringify(request));
    var sessionAttributes={};
    var filledSlots = delegateSlot.delegateSlotCollection(request, sessionAttributes, callback);

    //compose speechOutput that simply reads all the collected slot values
    var speechOutput = "AssignTask now";

    //Now let's recap the trip slot
    var recipient=request.intent.slots.recipient.value;
    var thing=request.intent.slots.thing.value;
    var duedate=request.intent.slots.duedate.value;

      speechOutput+= " information:"  + " recipient: "+ recipient + " thing: "+ thing + " duedate: " + duedate;

      console.log('session: '+JSON.stringify(session));

      //say the results
      callback(sessionAttributes, response.buildSpeechletResponse("AssignTask status", speechOutput, "", true));
}

exports.AssignTaskIntent = AssignTaskIntent;
