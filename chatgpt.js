class ChatGPT {

  generateConversationId() {
    return `conv-${Utilities.getUuid()}`;
  }

  request(promptObj, deal, conversationId = null, userCorrections = null) {
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!apiKey) {
        throw new Error('The API key was not found. Please set it in the script properties.');
    }

    const url = 'https://api.openai.com/v1/chat/completions';

    if(!conversationId){
      conversationId = this.generateConversationId();
    }
    if(!userCorrections){
      userCorrections = [];
    }
  
    const cache = CacheService.getScriptCache();
    let messages = JSON.parse(cache.get(conversationId)) || [];

    if (messages.length === 0) {
        messages.push({ role: "system", content: promptObj});
        messages.push({ role: "user", content: `${deal}` });
    } else if (userCorrections.length > 0) {
        messages.push({ role: "user", content: `Corrige el resultado anterior: ${userCorrections.join("\n")}` });
    }

    const payload = {
        model: GPT_MODEL,
        messages: messages,
        temperature: GPT_TEMPERATURE
    };

    const options = {
        method: "post",
        headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${apiKey}`
        },
        payload: JSON.stringify(payload)
    };

    try {
        const response = UrlFetchApp.fetch(url, options);
        const json = JSON.parse(response.getContentText());
        const usage = json.usage;
        const reply = json.choices[0].message.content.trim();
      
        messages.push({ role: "assistant", content: reply });

        cache.put(conversationId, JSON.stringify(messages), 300);
      
        return {
            "copy": reply,
            "tokens": usage['total_tokens'],
            "conversation_id": conversationId
        };
    } catch (error) {
        console.log(`Error: ${error.message}`);
    }
  }
}

