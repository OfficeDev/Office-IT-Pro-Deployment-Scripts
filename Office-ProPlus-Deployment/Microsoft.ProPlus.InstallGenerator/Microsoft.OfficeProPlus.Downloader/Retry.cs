using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;


    public class Retry
    {

        public static async Task<T> BlockAsync<T>(int retries, int secsDelay, Func<Task<T>> retryBock, CancellationToken token = new CancellationToken(), List<Exception> nonRetryExceptions = null, bool overrideDoNotRetry = false)
        {
            if (nonRetryExceptions == null) nonRetryExceptions = new List<Exception>();

            var backOff = new BackOff();
            var backOffStarted = false;

            while (true)
            {
                var useBackOff = false;
                try
                {
                    return await retryBock();
                }
                catch (Exception ex)
                {
                    if (token.IsCancellationRequested)
                    {
                        throw;
                    }

                    try
                    {
                        Trace.WriteLine("ERROR: " + ex.Message);
                        if (ex.InnerException != null)
                        {
                            Trace.WriteLine("ERROR: " + ex.InnerException.Message);
                        }
                    }
                    catch { }

                    if (ex.Message.ToLower().Contains("too many"))
                    {
                        useBackOff = true;
                        if (!backOffStarted && retries < 20) retries = 20;
                        backOffStarted = true;
                    }

                    LastErrorMessage = ex.ToString();
                    if (nonRetryExceptions.Any(exception => ex.GetType() == exception.GetType())) throw;
                    if (--retries < 0) throw;
                    if (DoNotRetry && !overrideDoNotRetry) throw;
                }

                RetryCount += 1;

                if (useBackOff)
                {
                    await backOff.RunAsync();
                }
                else
                {
                    await Task.Delay(secsDelay * 1000, token);
                }

            }
        }

        public static async Task BlockAsync(int retries, int secsDelay, Func<Task> retryBock, CancellationToken token = new CancellationToken(), List<string> nonRetryErrors = null, List<object> nonRetryExceptions = null, bool overrideDoNotRetry = false)
        {
            if (nonRetryExceptions == null) nonRetryExceptions = new List<object>();

            var backOff = new BackOff();
            var backOffStarted = false;

            while (true)
            {
                var useBackOff = false;
                try
                {
                    await retryBock();
                    break;
                }
                catch (Exception ex)
                {
                    if (token.IsCancellationRequested)
                    {
                        throw;
                    }

                    if (nonRetryErrors != null)
                    {
                        if (nonRetryErrors.Any(strError => ex.Message.Replace(" ", "").ToLower().Contains(strError.ToLower())))
                        {
                            break;
                        }
                    }

                    try
                    {
                        Trace.WriteLine("ERROR: " + ex.Message);
                        if (ex.InnerException != null)
                        {
                            Trace.WriteLine("ERROR: " + ex.InnerException.Message);
                        }
                    }
                    catch { }

                    if (ex.Message.ToLower().Contains("too many"))
                    {
                        useBackOff = true;
                        if (!backOffStarted && retries < 20) retries = 20;
                        backOffStarted = true;
                    }

                    LastErrorMessage = ex.ToString();
                    if (nonRetryExceptions.Any(exception => ex.GetType() == exception.GetType())) throw;
                    if (--retries < 0) throw;
                    if (DoNotRetry && !overrideDoNotRetry) throw;
                }

                RetryCount += 1;
                if (StopRetries) break;

                if (useBackOff)
                {
                    await backOff.RunAsync();
                }
                else
                {
                    await Task.Delay(secsDelay * 1000, token);
                }

                if (StopRetries) break;
            }
        }

        public static void Block(int retries, int secsDelay, Action retryBock, List<string> nonRetryErrors = null, List<object> nonRetryExceptions = null, bool overrideDoNotRetry = false)
        {
            if (nonRetryExceptions == null) nonRetryExceptions = new List<object>();

            var backOff = new BackOff();
            var backOffStarted = false;

            while (true)
            {
                var useBackOff = false;
                try
                {
                    retryBock();
                    break;
                }
                catch (Exception ex)
                {
                    if (ex.Message.ToLower().Contains("too many requests received"))
                    {
                        useBackOff = true;
                        if (!backOffStarted && retries < 20) retries = 20;
                        backOffStarted = true;
                    }

                    if (nonRetryErrors != null)
                    {
                        if (nonRetryErrors.Any(strError => ex.Message.Replace(" ", "").ToLower().Contains(strError.ToLower())))
                        {
                            break;
                        }
                    }

                    try
                    {
                        Trace.WriteLine("ERROR: " + ex.Message);
                        if (ex.InnerException != null)
                        {
                            Trace.WriteLine("ERROR: " + ex.InnerException.Message);
                        }
                    }
                    catch { }

                    LastErrorMessage = ex.ToString();
                    if (nonRetryExceptions.Any(exception => ex.GetType() == exception.GetType())) throw;
                    if (--retries < 0) throw;
                    if (DoNotRetry && !overrideDoNotRetry) throw;
                }
                RetryCount += 1;
                if (StopRetries) break;

                if (useBackOff)
                {
                    backOff.Run();
                }
                else
                {
                    System.Threading.Thread.Sleep(secsDelay * 1000);
                }

                if (StopRetries) break;
            }
        }

        public static T Block<T>(int retries, int secsDelay, Func<T> retryBock, List<Exception> nonRetryExceptions = null, bool overrideDoNotRetry = false)
        {
            if (nonRetryExceptions == null) nonRetryExceptions = new List<Exception>();

            var backOff = new BackOff();
            var backOffStarted = false;

            while (true)
            {
                var useBackOff = false;
                try
                {
                    return retryBock();
                }
                catch (Exception ex)
                {
                    if (ex.Message.ToLower().Contains("too many requests received"))
                    {
                        useBackOff = true;
                        if (!backOffStarted && retries < 20) retries = 20;
                        backOffStarted = true;
                    }

                    try
                    {
                        Trace.WriteLine("ERROR: " + ex.Message);
                        if (ex.InnerException != null)
                        {
                            Trace.WriteLine("ERROR: " + ex.InnerException.Message);
                        }
                    }
                    catch { }

                    RetryCount += 1;
                    LastErrorMessage = ex.ToString();
                    if (nonRetryExceptions.Any(exception => ex.GetType() == exception.GetType())) throw;
                    if (--retries < 0) throw;
                    if (DoNotRetry && !overrideDoNotRetry) throw;
                }
                RetryCount += 1;
                if (useBackOff)
                {
                    backOff.Run();
                }
                else
                {
                    System.Threading.Thread.Sleep(secsDelay * 1000);
                }
            }
        }

        public static bool DoNotRetry = false;
        public static bool StopRetries = false;

        public static int RetryCount = 0;
        public static string LastErrorMessage = "";

    }

