const axios = require('axios');

// ==========================================================
//             API KEY Pulled From User Input
// ==========================================================
var API_KEY = 'ENTER_API_KEY_HERE'; 

  async function sendMessageToChatGPT(message) {
    try {
      let token_count =  message.split(" ").length;
      const params = {
        prompt: message,
        model: 'text-davinci-003',
        temperature: 0.5,
        max_tokens: 4000-token_count-1 // Change this to say 4000-token_count-1 during demos
      };
      console.log(`Sending Message`);
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
      throw new Error(`HTTP error, status = ${response.status}`);
    }
  } catch (error) {
    if (error.response) {
      if (error.response.status === 400) {
        console.log("Error: The request was invalid or malformed. Try Again");
        return ["Error: The request was invalid or malformed. Try Again", 400];
      }
      else if (error.response.status === 401) {
        console.log("Error: Invalid API key or not authorized to access the API");
        return ["Error: Invalid API key or not authorized to access the API", 401];
      } 
      else if (error.response.status === 403) {
        console.log("Error: The API key does not have the necessary permissions for the request. Email: dennis.cherian@hellotinker.com for a new access key");
        return ["Error: The API key does not have the necessary permissions for the request. Email: dennis.cherian@hellotinker.com for a new access key", 403];
      }
      else if (error.response.status === 429) {
        console.log("Error: The user has sent too many requests in a given amount of time. Please wait a minute and try again");
        return ["Error: The user has sent too many requests in a given amount of time. Please wait a minute and try again", 429];
      }
      else if (error.response.status === 500) {
        console.log("Error: The server failed to generate a reply");
        return ["Error: The server failed to generate a reply", 500];
      }
      else {
        console.log("Error: " + error.message);
        return ["Error: The server failed to generate a reply", 500];
      }
    } else {
      console.log("Error: " + error.message);
      return ["Error: The server failed to generate a reply", 500];
    }
  }
}


  export async function getMessageFromChatGPT(message) {
      const [response, status] = await sendMessageToChatGPT(message);
      console.log("Got formatted server response");
      console.log(response, status);
      return [response, status];
  }

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("submit-chat").onclick = submitHandler;
  }
});

export async function submitHandler(event) {
  return Word.run(async (context) => {
    document.getElementById("error_pane").innerText = "";
    document.getElementById("submit-chat").disabled = true;
    document.getElementById("loading-animation").style.visibility = 'visible';
    API_KEY = document.getElementById("key-for-chat-gpt").value;
    var message = document.getElementById("text-input-for-chat-gpt").value;
    let [reply, status] = await getMessageFromChatGPT(message);
    console.log("In main again");
    if (status === 200){
      let reply_lines = reply.split("\n")
      for (let i = 0; i < reply_lines.length; i++) {
        context.document.body.insertParagraph(reply_lines[i], Word.InsertLocation.end);
      }
      await context.sync();
      console.log("Working");
    }
    else {
      console.log("Post error");
      document.getElementById("error_pane").innerText = reply;
    }

    document.getElementById("loading-animation").style.visibility = 'hidden';
    document.getElementById("submit-chat").disabled = false;
  });
}
