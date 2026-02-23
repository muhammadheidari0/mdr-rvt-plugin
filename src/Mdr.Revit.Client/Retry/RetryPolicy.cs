using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace Mdr.Revit.Client.Retry
{
    public sealed class RetryPolicy
    {
        public RetryPolicy(int maxAttempts = 3, TimeSpan? delay = null)
        {
            if (maxAttempts <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(maxAttempts), "Max attempts must be greater than zero.");
            }

            MaxAttempts = maxAttempts;
            Delay = delay ?? TimeSpan.FromMilliseconds(400);
        }

        public int MaxAttempts { get; }

        public TimeSpan Delay { get; }

        public async Task<T> ExecuteAsync<T>(
            Func<CancellationToken, Task<T>> operation,
            CancellationToken cancellationToken)
        {
            if (operation == null)
            {
                throw new ArgumentNullException(nameof(operation));
            }

            int attempt = 0;
            Exception? lastException = null;

            while (attempt < MaxAttempts)
            {
                attempt++;

                try
                {
                    return await operation(cancellationToken).ConfigureAwait(false);
                }
                catch (Exception ex) when (IsTransient(ex) && attempt < MaxAttempts)
                {
                    lastException = ex;
                    await Task.Delay(Delay, cancellationToken).ConfigureAwait(false);
                }
            }

            throw lastException ?? new InvalidOperationException("Retry operation failed without exception details.");
        }

        private static bool IsTransient(Exception exception)
        {
            if (exception is HttpRequestException)
            {
                return true;
            }

            if (exception is TaskCanceledException)
            {
                return true;
            }

            return false;
        }
    }
}
