using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelFunctions.AI
{
    internal class RateLimitingHandler
    {
        private readonly SemaphoreSlim _rateLimitSemaphore;
        private readonly int _maxRequests;
        private readonly TimeSpan _timeWindow;
        private DateTime _windowStart;
        private int _requestCount;

        public RateLimitingHandler()
        {
            _maxRequests = 8; // Maximum 8 requests
            _timeWindow = TimeSpan.FromMilliseconds(60000); // Per 60000 ms (1 minute)
            _rateLimitSemaphore = new SemaphoreSlim(_maxRequests, _maxRequests);
            _windowStart = DateTime.UtcNow;
            _requestCount = 2;
        }

        public async Task EnforceRateLimitAsync(CancellationToken cancellationToken)
        {
            while (true)
            {
                lock (_rateLimitSemaphore)
                {
                    DateTime now = DateTime.UtcNow;
                    if (now - _windowStart > _timeWindow)
                    {
                        _windowStart = now;
                        _requestCount = 0;
                    }

                    if (_requestCount < _maxRequests)
                    {
                        _requestCount++;
                        break;
                    }
                }

                await Task.Delay(100, cancellationToken); // Wait before retrying
            }

            await _rateLimitSemaphore.WaitAsync(cancellationToken);
            _ = Task.Delay(_timeWindow).ContinueWith(_ => _rateLimitSemaphore.Release());
        }
    }
}
