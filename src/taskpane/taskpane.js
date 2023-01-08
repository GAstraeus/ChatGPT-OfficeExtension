const axios = require('axios');

// ==========================================================
//             API KEY Pulled From User Input
// ==========================================================
var API_KEY = 'ENTER_API_KEY_HERE'; 

  async function sendMessageToChatGPT(message) {
    try {
      // Set the prompt for GPT-3 to respond to
      let token_count =  message.split(" ").length;
      // Set the parameters for the GPT-3 request
      const params = {
        prompt: message,
        model: 'text-davinci-003',
        temperature: 0.5,
        max_tokens: 4000-token_count-1 // Change this to say 4000-token_count-1 during demos
      };
      // Call the GPT-3 API
      console.log(`Sending Message`);
      const response = await axios.post('https://api.openai.com/v1/completions', params, {
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${API_KEY}`
        }
      });
      console.log(response.data.choices[0].text);
      return await response;
    } catch (error) {
      console.error(error);
      return null;
    }
  }

  export async function getMessageFromChatGPT(message) {
    const response = await sendMessageToChatGPT(message);
    if (response == null) {
      console.log("Error calling chat gpt");
      return "";
    }
    return response.data.choices[0].text;
  }

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("submit-chat").onclick = submitHandler;
  }
});

export async function submitHandler(event) {
  return Word.run(async (context) => {
    document.getElementById("submit-chat").disabled = true;
    document.getElementById("loading-animation").style.visibility = 'visible';
    API_KEY = document.getElementById("key-for-chat-gpt").value;
    var message = document.getElementById("text-input-for-chat-gpt").value;
    let reply = await getMessageFromChatGPT(message);

    let reply_lines = reply.split("\n")
    for (let i = 0; i < reply_lines.length; i++) {
      context.document.body.insertParagraph(reply_lines[i], Word.InsertLocation.end);
    }

    await context.sync();
    document.getElementById("loading-animation").style.visibility = 'hidden';
    document.getElementById("submit-chat").disabled = false;
  });
}
