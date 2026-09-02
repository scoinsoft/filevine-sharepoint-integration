const { AsyncLocalStorage } = require('async_hooks');

const requestContext = new AsyncLocalStorage();

function isServerless() {
  return Boolean(process.env.VERCEL || process.env.AWS_LAMBDA_FUNCTION_NAME);
}

function isVercel() {
  return Boolean(process.env.VERCEL);
}

function getFunctionBudgetMs() {
  const configured = Number(process.env.FUNCTION_MAX_DURATION_MS);
  if (Number.isFinite(configured) && configured > 0) {
    return Math.floor(configured);
  }
  if (!isServerless()) {
    return 24 * 60 * 60 * 1000;
  }
  // Stay under Vercel Pro maxDuration (300s). Hobby caps at 60s at the platform.
  return 270000;
}

function getStopBufferMs() {
  const configured = Number(process.env.FUNCTION_STOP_BUFFER_MS);
  if (Number.isFinite(configured) && configured > 0) {
    return Math.floor(configured);
  }
  return isServerless() ? 45000 : 0;
}

function runWithRequestContext(store, fn) {
  return requestContext.run(store, fn);
}

function startRequestDeadline() {
  const deadlineAt = Date.now() + getFunctionBudgetMs();
  return { deadlineAt, startedAt: Date.now() };
}

function remainingMs() {
  const store = requestContext.getStore();
  if (!store?.deadlineAt) {
    return Number.POSITIVE_INFINITY;
  }
  return store.deadlineAt - Date.now();
}

function shouldStopForDeadline() {
  if (!isServerless()) {
    return false;
  }
  return remainingMs() <= getStopBufferMs();
}

module.exports = {
  isServerless,
  isVercel,
  getFunctionBudgetMs,
  getStopBufferMs,
  runWithRequestContext,
  startRequestDeadline,
  remainingMs,
  shouldStopForDeadline,
};
