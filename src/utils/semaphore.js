class Semaphore {
  constructor(max) {
    this.max = Math.max(1, max);
    this.inUse = 0;
    this.queue = [];
  }

  acquire() {
    if (this.inUse < this.max) {
      this.inUse += 1;
      return Promise.resolve();
    }

    return new Promise((resolve) => {
      this.queue.push(resolve);
    });
  }

  release() {
    if (this.queue.length > 0) {
      const next = this.queue.shift();
      next();
      return;
    }
    this.inUse = Math.max(0, this.inUse - 1);
  }

  async run(fn) {
    await this.acquire();
    try {
      return await fn();
    } finally {
      this.release();
    }
  }
}

module.exports = { Semaphore };
