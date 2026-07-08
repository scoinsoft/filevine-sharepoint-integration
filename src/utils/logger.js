function log(message, data) {
  const timestamp = new Date().toISOString();
  if (data !== undefined) {
    console.log(`[${timestamp}] ${message}`, data);
  } else {
    console.log(`[${timestamp}] ${message}`);
  }
}

function logError(message, error) {
  const timestamp = new Date().toISOString();
  console.error(`[${timestamp}] ${message}`, error?.message || error);
  if (error?.response?.data) {
    console.error(`[${timestamp}] Response data:`, error.response.data);
  }
}

module.exports = { log, logError };
