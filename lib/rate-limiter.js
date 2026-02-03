/**
 * Rate Limiter with Exponential Backoff
 * Prevents hitting Microsoft Graph API rate limits
 */

class RateLimiter {
  constructor(options = {}) {
    this.minDelay = options.minDelay || 100; // Minimum delay between requests (ms)
    this.maxDelay = options.maxDelay || 30000; // Maximum delay (30 seconds)
    this.maxRetries = options.maxRetries || 3;
    this.lastRequestTime = 0;
    this.consecutiveErrors = 0;
  }

  /**
   * Wait before making a request
   */
  async waitForSlot() {
    const now = Date.now();
    const timeSinceLastRequest = now - this.lastRequestTime;
    const delay = Math.max(0, this.minDelay - timeSinceLastRequest);
    
    if (delay > 0) {
      await this.sleep(delay);
    }
    
    this.lastRequestTime = Date.now();
  }

  /**
   * Calculate backoff delay for retries
   * @param {number} attempt - The current attempt number (0-indexed)
   * @returns {number} - Delay in milliseconds
   */
  getBackoffDelay(attempt) {
    // Exponential backoff: 1s, 2s, 4s, 8s, etc.
    const delay = Math.min(
      this.maxDelay,
      Math.pow(2, attempt) * 1000 + Math.random() * 1000 // Add jitter
    );
    return delay;
  }

  /**
   * Execute a function with rate limiting and retry logic
   * @param {Function} fn - The async function to execute
   * @returns {Promise<any>} - The result of the function
   */
  async execute(fn) {
    let lastError;
    
    for (let attempt = 0; attempt <= this.maxRetries; attempt++) {
      try {
        await this.waitForSlot();
        const result = await fn();
        this.consecutiveErrors = 0;
        return result;
      } catch (error) {
        lastError = error;
        
        // Check if it's a rate limit error (429)
        const isRateLimited = error.message?.includes('429') || 
                             error.statusCode === 429 ||
                             error.code === 'TooManyRequests';
        
        // Check if it's a retryable error
        const isRetryable = isRateLimited ||
                           error.message?.includes('503') ||
                           error.message?.includes('504') ||
                           error.code === 'ServiceUnavailable';
        
        if (!isRetryable || attempt === this.maxRetries) {
          throw error;
        }
        
        this.consecutiveErrors++;
        const delay = this.getBackoffDelay(attempt);
        console.error(`Rate limited or temporary error. Retrying in ${Math.round(delay/1000)}s... (attempt ${attempt + 1}/${this.maxRetries})`);
        await this.sleep(delay);
      }
    }
    
    throw lastError;
  }

  /**
   * Sleep for a specified duration
   * @param {number} ms - Duration in milliseconds
   */
  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

// Export a singleton instance
export const rateLimiter = new RateLimiter();

// Export the class for custom instances
export { RateLimiter };
