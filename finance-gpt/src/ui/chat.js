// Helper function to add a message to the chat
function addMessageToChat(role, content) {
  const chatHistory = document.getElementById("chatHistory");
  const messageDiv = document.createElement('div');
  messageDiv.className = `chat-message ${role}-message`;
  
  const header = document.createElement('div');
  header.className = 'message-header';
  
  const roleSpan = document.createElement('span');
  roleSpan.className = 'message-role';
  roleSpan.textContent = role === 'user' ? 'You' : 'Assistant';
  header.appendChild(roleSpan);
  
  const messageContent = document.createElement('div');
  messageContent.className = 'message-content';
  messageContent.innerHTML = content;
  
  messageDiv.appendChild(header);
  messageDiv.appendChild(messageContent);
  chatHistory.appendChild(messageDiv);
  
  return messageDiv;
}

// Helper function to format the response with syntax highlighting
function formatResponse(response) {
  return response.replace(
    /```javascript([\s\S]*?)```/g,
    (match, code) => `<code class="javascript">${code.trim()}</code>`
  );
}

function createStatusMessage(message, type) {
  const statusDiv = document.createElement("div");
  statusDiv.className = `status-message ${type}`;
  statusDiv.textContent = message;
  return statusDiv;
}

export { addMessageToChat, formatResponse, createStatusMessage }; 