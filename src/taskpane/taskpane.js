import { errorMessage400, errorMessage401, errorMessage403, errorMessage429, noMessageInput, noAPIKeyInput } from './common/error_constants';

const axios = require('axios');

// ==========================================================
//             API KEY Pulled From User Input
// ==========================================================
var API_KEY = 'ENTER_API_KEY_HERE'; 

function invalidInput(s)
{
    return (s == null || s == "");
}

async function sendMessageToServer(message) {
  try {
    let token_count =  message.split(" ").length;
    const params = {
      prompt: message,
      model: 'text-davinci-003',
      temperature: 0.5,
      max_tokens: 4000-token_count-10
    };
    console.log(`Sending to Process`);
    const response = await axios.post('https://api.openai.com/v1/completions', params, {
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${API_KEY}`
      }
    });
  if (response.status === 200) {
    // The request was successful
      return [response.data.choices[0].text, 200];
  } else {
      console.log("Unexpected Error");
      return [errorMessage500, 500];
  }
} catch (error) {
  if (error.response) {
    if (error.response.status === 400) {
      console.log(errorMessage400);
      return [errorMessage400, error.response.status];
    }
    else if (error.response.status === 401) {
      console.log(errorMessage401);
      return [errorMessage401, error.response.status];
    } 
    else if (error.response.status === 403) {
      console.log(errorMessage403);
      return [errorMessage403, error.response.status];
    }
    else if (error.response.status === 429) {
      console.log(errorMessage429);
      return [errorMessage429, error.response.status];
    }
    else if (error.response.status === 500) {
      console.log(errorMessage500);
      return [errorMessage500, error.response.status];
    }
    else {
      console.log("Error: " + error.message);
      return [errorMessage500, 500];
    }
  } else {
    console.log("Error: " + error.message);
    return [errorMessage500, 500];
  }
}
}


export async function getMessageFromServer(message) {
    const [response, status] = await sendMessageToServer(message);
    console.log("Response Recieved", response, status);
    return [response, status];
}

Office.onReady((info) => {
if (info.host === Office.HostType.Word) {
  document.getElementById("submit-chat").onclick = submitHandler;
  document.getElementById("error_pane").innerText = "";
}
});

export async function submitHandler(event) {
return Word.run(async (context) => {
  document.getElementById("error_pane").innerText = "";
  document.getElementById("submit-chat").disabled = true;
  document.getElementById("loading-animation").style.visibility = 'visible';
  API_KEY = document.getElementById("key-for-api").value;
  var message = document.getElementById("text-input-for-api").value;
  
  var reply = "";
  var status = "";
  if (invalidInput(API_KEY)){
    reply = noAPIKeyInput;
    status = 401;
  }
  else if (invalidInput(message)){
    reply = noMessageInput;
    status = 401;
  }
  else {
    [reply, status] = await getMessageFromServer(message);
  }

  if (status === 200){
    let reply_lines = reply.split("\n")
    for (let i = 0; i < reply_lines.length; i++) {
      context.document.body.insertParagraph(reply_lines[i], Word.InsertLocation.end);
    }
    await context.sync();
  }
  else {
    document.getElementById("error_pane").innerText = reply;
  }

  document.getElementById("loading-animation").style.visibility = 'hidden';
  document.getElementById("submit-chat").disabled = false;
});
}
