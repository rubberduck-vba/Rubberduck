using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.Client.Abstract;
using Rubberduck.Settings;

namespace Rubberduck.VersionCheck
{
    public interface IPublicApiClient
    {
        Task<IEnumerable<Tag>> GetLatestTagsAsync(CancellationToken token);
    }

    public class PublicApiClient : ApiClientBase, IPublicApiClient
    {
        private static readonly string PublicTagsEndPoint = "public/tags";

        public PublicApiClient(IGeneralSettings settings, IHttpClientProvider clientProvider) 
            : base(settings, clientProvider)
        {
        }

        public async Task<IEnumerable<Tag>> GetLatestTagsAsync(CancellationToken token)
        {
            return await GetResponse<Tag[]>(PublicTagsEndPoint, token);
        }
    }
}
