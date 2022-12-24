const axios = require('axios');

// ==========================================================
//                  ENTER YOUR API KEY HERE 
// ==========================================================
const API_KEY = 'ENTER_API_KEY_HERE'; 

async function initializeChatGPT() { //THIS INITIALIZATION WONT BE REQUIRED IF USING CUSTOM MODEL
    // Set the prompt for GPT-3 to respond to
    // const seed = 'The NYC personal injury law firm of David Resnick & Associates, P.C. provides professional and caring legal assistance to victims of injury and negligence in the New York City area. \nDavid Resnick founded the firm in 1998 after working in large law firms where he saw a need for greater client communication and more personal care. He wanted to help everyday folks who have had the misfortune to be injured in an accident. His early experiences at other firms taught him that personal care and attention, as well as communication with clients, should be the cornerstones of a law practice – principles that he found lacking elsewhere. \nHis effort to change the way attorneys relate to their clients started when he launched his own legal practice from a tiny office in his apartment in the city. His wife, a student at the time, helped out by answering phones in between classes. \nDavid Resnick & Associates, P.C. has grown into a successful New York personal injury law firm in midtown Manhattan. Although the firm has grown, Resnick’s dedication to client relationships and his core values remain the same. \nConstant communication is key at David Resnick & Associates, P.C. He frequently visits clients in their homes or at the hospital in order to fully understand the trauma that an accident has caused. Meeting a client’s family is something that David Resnick believes is an important part of many injury cases because when one member of a family is injured, other family members often suffer as well. David Resnick is devoted to his own family, and he understands that clients and their loved ones need help to rebuild their lives after a traumatic accident. \nWhat Our Clients Say \nDavid Resnick & Associates were very professional and helpful in my time of need. Very sympathetic to my situation. The end result was more then I expected and I’m more then satisfied. I highly recommend David Resnick & Associates for the best representation. On a personal note I felt comfortable... \n- Carlos \nAt David Resnick & Associates, P.C., we understand how the victim of an accident feels and we also understand that as a victim, you are essentially putting your future in our hands when you retain our firm to represent you.';
    const seed = 'The NYC personal injury law firm of David Resnick & Associates, P.C. provides professional and caring legal assistance to victims of injury and negligence in the New York City area';
    await sendMessageToChatGPT(seed);
}

  async function sendMessageToChatGPT(message) { //THIS INITIALIZATION WONT BE REQUIRED IF USING CUSTOM MODEL
    try {
      // Set the prompt for GPT-3 to respond to
      let token_count =  message.split(" ").length;
      // Set the parameters for the GPT-3 request
      const params = {
        prompt: message,
        model: 'text-davinci-003',
        temperature: 0.5,
        max_tokens: 2000-token_count+2
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
  initializeChatGPT();
});

export async function submitHandler(event) {
  return Word.run(async (context) => {
    var message = document.getElementById("text-input-for-chat-gpt").value;
    let reply = await getMessageFromChatGPT(message);

    let reply_lines = reply.split("\n")
    for (let i = 0; i < reply_lines.length; i++) {
      context.document.body.insertParagraph(reply_lines[i], Word.InsertLocation.end);
    }
    await context.sync();
  });
}
