using Newtonsoft.Json;
using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Rubberduck.Client.Abstract
{
    public abstract class ApiClientBase : IDisposable
    {
        protected static readonly string UserAgentName = "Rubberduck";
        protected static readonly string BaseUrl = "https://api.rubberduckvba.com/api/v1/";
        protected static readonly string ContentTypeApplicationJson = "application/json";
        protected static readonly int MaxAttempts = 3;
        protected static readonly TimeSpan RetryCooldownDelay = TimeSpan.FromSeconds(1);

        protected readonly Lazy<HttpClient> _client;

        protected ApiClientBase()
        {
            _client = new Lazy<HttpClient>(() => GetClient());
        }

        protected HttpClient GetClient()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            var client = new HttpClient();
            return ConfigureClient(client);
        }

        protected virtual HttpClient ConfigureClient(HttpClient client)
        {
            var userAgentVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString(3);
            var userAgentHeader = new ProductInfoHeaderValue(UserAgentName, userAgentVersion);

            client.DefaultRequestHeaders.UserAgent.Add(userAgentHeader);
            return client;
        }

        protected virtual async Task<TResult> GetResponse<TResult>(string route, CancellationToken? cancellationToken = null)
        {
            var uri = new Uri($"{BaseUrl}{route}");

            var attempt = 0;
            var token = cancellationToken ?? CancellationToken.None;

            while (!token.IsCancellationRequested && attempt <= MaxAttempts)
            {
                attempt++;
                var delay = attempt == 0 ? TimeSpan.Zero : RetryCooldownDelay;

                var (success, result) = await TryGetResponse<TResult>(uri, delay, token);
                if (success)
                {
                    return result;
                }
            }

            token.ThrowIfCancellationRequested();
            throw new InvalidOperationException($"API call failed to return a result after {attempt} attempts.");
        }

        private async Task<(bool, TResult)> TryGetResponse<TResult>(Uri uri, TimeSpan delay, CancellationToken token)
        {
            if (delay != TimeSpan.Zero)
            {
                await Task.Delay(delay);
            }

            token.ThrowIfCancellationRequested();

            try
            {
                using (var client = GetClient())
                {
                    using (var response = await client.GetAsync(uri))
                    {
                        response.EnsureSuccessStatusCode();
                        token.ThrowIfCancellationRequested();

                        var content = await response.Content.ReadAsStringAsync();
                        var result = JsonConvert.DeserializeObject<TResult>(content);

                        return (true, result);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch
            {
                return default;
            }
        }

        protected virtual async Task<T> Post<T>(string route, T args, CancellationToken? cancellationToken = null) => await Post<T, T>(route, args, cancellationToken ?? CancellationToken.None);

        protected virtual async Task<TResult> Post<TArgs, TResult>(string route, TArgs args, CancellationToken? cancellationToken = null)
        {
            var uri = new Uri($"{BaseUrl}{route}");
            string json;
            try
            {
                json = JsonConvert.SerializeObject(args);
            }
            catch (Exception exception)
            {
                throw new ArgumentException("The specified arguments could not be serialized.", exception);
            }

            var attempt = 0;
            var token = cancellationToken ?? CancellationToken.None;

            while (!token.IsCancellationRequested && attempt <= MaxAttempts)
            {
                attempt++;
                var delay = attempt == 0 ? TimeSpan.Zero : RetryCooldownDelay;

                var (success, result) = await TryPost<TResult>(uri, json, delay, token);
                if (success)
                {
                    return result;
                }
            }

            token.ThrowIfCancellationRequested();
            throw new InvalidOperationException($"API call failed to return a result after {attempt} attempts.");
        }

        private async Task<(bool, TResult)> TryPost<TResult>(Uri uri, string body, TimeSpan delay, CancellationToken token)
        {
            if (delay != TimeSpan.Zero)
            {
                await Task.Delay(delay);
            }

            token.ThrowIfCancellationRequested();

            try
            {
                using (var client = GetClient())
                {
                    var content = new StringContent(body, Encoding.UTF8, ContentTypeApplicationJson);
                    using (var response = await client.PostAsync(uri, content, token))
                    {
                        response.EnsureSuccessStatusCode();
                        token.ThrowIfCancellationRequested();

                        var jsonResult = await response.Content.ReadAsStringAsync();
                        var result = JsonConvert.DeserializeObject<TResult>(jsonResult);

                        return (true, result);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch
            {
                return default;
            }
        }

        public void Dispose()
        {
            if (_client.IsValueCreated)
            {
                _client.Value.Dispose();
            }
        }
    }
}