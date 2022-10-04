using System;
using System.Configuration;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace Rubberduck.Core.WebApi
{
    public abstract class ApiClientBase
    {
        public const string ContentTypeApplicationJson = "application/json";

        private readonly string _baseUrl;
        protected TimeSpan GetRequestTimeout { get; }
        protected TimeSpan PostRequestTimeout { get; }

        private static readonly ProductInfoHeaderValue UserAgent =
            new ProductInfoHeaderValue("Rubberduck", System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString());

        protected ApiClientBase()
        {
            _baseUrl = Properties.Settings.Default.WebApiBaseUrl;
        }

        protected string BaseUrl => _baseUrl;

        protected virtual HttpClient GetClient(string contentType = ContentTypeApplicationJson)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.UserAgent.Add(UserAgent);
            return client;
        }

        protected virtual async Task<TResult> GetResponseAsync<TResult>(string route)
        {
            var uri = new Uri($"{_baseUrl}/{route}");
            try
            {
                using (var client = GetClient())
                {
                    //client.Timeout = GetRequestTimeout;
                    using (var response = await client.GetAsync(uri))
                    {
                        response.EnsureSuccessStatusCode();
                        var content = await response.Content.ReadAsStringAsync();

                        return JsonSerializer.Deserialize<TResult>(content, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                    }
                }
            }
            catch (Exception exception)
            {
                throw new ApiException(exception);
            }
        }

        protected virtual async Task<T> PostAsync<T>(string route, T args) => await PostAsync<T, T>(route, args);

        protected virtual async Task<TResult> PostAsync<TArgs, TResult>(string route, TArgs args)
        {
            var uri = new Uri($"{_baseUrl}/{route}");
            string json;
            try
            {
                json = JsonSerializer.Serialize(args);
            }
            catch (Exception exception)
            {
                throw new ArgumentException("The specified arguments could not be serialized.", exception);
            }

            try
            {
                using (var client = GetClient())
                {
                    //client.Timeout = PostRequestTimeout;
                    using (var response = await client.PostAsync(uri, new StringContent(json, Encoding.UTF8, ContentTypeApplicationJson)))
                    {
                        response.EnsureSuccessStatusCode();
                        var content = await response.Content.ReadAsStringAsync();
                        var result = JsonSerializer.Deserialize<TResult>(content, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                        return result;
                    }
                }
            }
            catch (Exception exception)
            {
                throw new ApiException(exception);
            }
        }
    }
}
