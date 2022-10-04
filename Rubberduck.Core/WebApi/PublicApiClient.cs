using Rubberduck.Core.WebApi.Model;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Rubberduck.Core.WebApi
{
    public class PublicApiClient : ApiClientBase, IPublicApiClient
    {
        private static readonly string FeatureEndpoint = "Feature";
        private static readonly string FeatureItemEndpoint = "FeatureItem";
        private static readonly string TagsEndpoint = "Tags";

        public async Task<IEnumerable<Tag>> GetLatestTagsAsync()
        {
            try
            {
                return await GetResponseAsync<Tag[]>(TagsEndpoint);
            }
            catch (ApiException)
            {
                return Enumerable.Empty<Tag>();
            }
        }

        public async Task<Feature> GetFeatureAsync(string name)
        {
            try
            {
                return await GetResponseAsync<Feature>($"{FeatureEndpoint}/{name}");
            }
            catch (ApiException)
            {
                return null;
            }
        }

        public async Task<FeatureItem> GetFeatureItemAsync(string name)
        {
            try
            {
                return await GetResponseAsync<FeatureItem>($"{FeatureItemEndpoint}/{name}");
            }
            catch (ApiException)
            {
                return null;
            }
        }
    }
}
