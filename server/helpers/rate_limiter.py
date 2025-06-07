from fastapi import HTTPException, status
from fastapi.responses import JSONResponse
import time
from typing import Dict, Tuple
import asyncio
from datetime import datetime, timedelta

# In-memory storage for rate limiting
# In production, you should use Redis or another distributed cache
rate_limit_store: Dict[str, Tuple[int, float]] = {}

async def rate_limit(key: str, max_requests: int = 5, window_seconds: int = 60) -> bool:
    """
    Rate limit requests based on a key.
    
    Args:
        key: The key to rate limit (e.g., IP address or user identifier)
        max_requests: Maximum number of requests allowed in the time window
        window_seconds: Time window in seconds
        
    Returns:
        bool: True if request is allowed, False if rate limited
    """
    current_time = time.time()
    
    # Get the current count and window start time for this key
    count, window_start = rate_limit_store.get(key, (0, current_time))
    
    # If the window has expired, reset the counter
    if current_time - window_start > window_seconds:
        count = 0
        window_start = current_time
    
    # If we've hit the rate limit, return False
    if count >= max_requests:
        return False
    
    # Increment the counter and update the store
    count += 1
    rate_limit_store[key] = (count, window_start)
    
    return True

def get_rate_limit_headers(key: str, max_requests: int = 5, window_seconds: int = 60) -> Dict[str, str]:
    """
    Get rate limit headers for the response.
    
    Args:
        key: The key being rate limited
        max_requests: Maximum number of requests allowed
        window_seconds: Time window in seconds
        
    Returns:
        Dict[str, str]: Headers to include in the response
    """
    count, window_start = rate_limit_store.get(key, (0, time.time()))
    reset_time = datetime.fromtimestamp(window_start + window_seconds)
    
    return {
        "X-RateLimit-Limit": str(max_requests),
        "X-RateLimit-Remaining": str(max(0, max_requests - count)),
        "X-RateLimit-Reset": reset_time.isoformat()
    }

async def check_rate_limit(request, key_prefix: str = "default"):
    """
    Check rate limit for a request and raise an exception if exceeded.
    
    Args:
        request: The FastAPI request object
        key_prefix: Prefix for the rate limit key
    """
    # Use IP address as the key for rate limiting
    client_ip = request.client.host
    key = f"{key_prefix}:{client_ip}"
    
    if not await rate_limit(key):
        headers = get_rate_limit_headers(key)
        raise HTTPException(
            status_code=status.HTTP_429_TOO_MANY_REQUESTS,
            detail="Too many requests. Please try again later.",
            headers=headers
        ) 