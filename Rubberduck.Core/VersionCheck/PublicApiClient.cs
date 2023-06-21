using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.Client.Abstract;

namespace Rubberduck.VersionCheck
{
    public class PublicApiClient : ApiClientBase
    {
        private static readonly string PublicTagsEndPoint = "public/tags";

        public async Task<IEnumerable<Tag>> GetLatestTagsAsync()
        {
            var tokenSource = new CancellationTokenSource();
            tokenSource.CancelAfter(TimeSpan.FromSeconds(5));

            return await GetResponse<Tag[]>(PublicTagsEndPoint, tokenSource.Token);
        }
    }
}
