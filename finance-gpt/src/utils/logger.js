// Centralized logging utility
export const Logger = {
  responseArea: null,

  setResponseArea(element) {
    this.responseArea = element;
  },

  updateUI(message, data = null) {
    if (this.responseArea) {
      // Format the message
      let formattedMessage = message;
      
      // If data is provided, format it nicely
      if (data) {
        if (typeof data === 'object') {
          // For objects, format them nicely with indentation
          formattedMessage += '\n' + JSON.stringify(data, null, 2)
            .split('\n')
            .map(line => '  ' + line)  // Add indentation
            .join('\n');
        } else {
          formattedMessage += ' ' + data;
        }
      }
      
      // Always add a newline at the end of each message
      this.responseArea.innerHTML += formattedMessage + '\n\n';
      // Auto-scroll to bottom
      this.responseArea.scrollTop = this.responseArea.scrollHeight;
    }
  },

  startOperation: (name) => {
    const message = `\n=== ${name} Started ===`;
    Logger.updateUI(message);
  },
  
  endOperation: (name) => {
    const message = `=== ${name} Completed ===`;
    Logger.updateUI(message);
  },
  
  error: (name, error, additionalInfo = {}) => {
    let message = `\n❌ Error in ${name}:`;
    message += `\n  Message: ${error.message}`;
    
    if (additionalInfo && Object.keys(additionalInfo).length > 0) {
      message += '\n  Details:';
      Logger.updateUI(message, additionalInfo);
    } else {
      Logger.updateUI(message);
    }
  },
  
  info: (message, data = null) => {
    // If it's a section header (starts with ===)
    if (message.startsWith('===')) {
      Logger.updateUI('\n' + message);
    } else {
      // For regular info messages
      Logger.updateUI('  ' + message, data);
    }
  },

  success: (message) => {
    Logger.updateUI('  ✓ ' + message);
  }
}; 