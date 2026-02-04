/**
 * Rate Limiter with Exponential Backoff
 * Prevents hitting Microsoft Graph API rate limits
 */

import type { RateLimiterOptions } from '../types.js';

interface RateLimitError extends Error {
  statusCode?: number;
  code?: string;
}

export class RateLimiter {
  private minDelay: number;
  private maxDelay: number;
  private maxRetries: number;
  private lastRequestTime: number;
  private consecutiveErrors: number;

  constructor(options: RateLimiterOptions = {}) {
    this.minDelay = options.minDelay ?? 100; // Minimum delay between requests (ms)
    this.maxDelay = options.maxDelay ?? 30000; // Maximum delay (30 seconds)
    this.maxRetries = options.maxRetries ?? 3;
    this.lastRequestTime = 0;
    this.consecutiveErrors = 0;
  }

  /**
   * Wait before making a request
   */
  private async waitForSlot(): Promise<void> {
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
   */
  private getBackoffDelay(attempt: number): number {
    // Exponential backoff: 1s, 2s, 4s, 8s, etc.
    const delay = Math.min(
      this.maxDelay,
      Math.pow(2, attempt) * 1000 + Math.random() * 1000 // Add jitter
    );
    return delay;
  }

  /**
   * Execute a function with rate limiting and retry logic
   */
  async execute<T>(fn: () => Promise<T>): Promise<T> {
    let lastError: Error | undefined;
    
    for (let attempt = 0; attempt <= this.maxRetries; attempt++) {
      try {
        await this.waitForSlot();
        const result = await fn();
        this.consecutiveErrors = 0;
        return result;
      } catch (error) {
        lastError = error instanceof Error ? error : new Error(String(error));
        const rateLimitError = error as RateLimitError;
        
        // Check if it's a rate limit error (429)
        const isRateLimited = rateLimitError.message?.includes('429') || 
                             rateLimitError.statusCode === 429 ||
                             rateLimitError.code === 'TooManyRequests';
        
        // Check if it's a retryable error
        const isRetryable = isRateLimited ||
                           rateLimitError.message?.includes('503') ||
                           rateLimitError.message?.includes('504') ||
                           rateLimitError.code === 'ServiceUnavailable';
        
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
   */
  private sleep(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

// Export a singleton instance
export const rateLimiter = new RateLimiter();
