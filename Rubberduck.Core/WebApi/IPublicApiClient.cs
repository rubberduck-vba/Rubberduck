using Rubberduck.Core.WebApi.Model;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Rubberduck.Core.WebApi
{
    /// <summary>
    /// An API client that provides website content through anonymous requests.
    /// </summary>
    public interface IPublicApiClient
    {
        /// <summary>
        /// Gets a feature with its sub-features and their respective feature items.
        /// </summary>
        /// <remarks>
        /// A separate request is required to retrieve a particular feature item's details.
        /// </remarks>
        /// <param name="name">Uniquely identifies the feature to get.</param>
        Task<Feature> GetFeatureAsync(string name);
        /// <summary>
        /// Gets the specified feature item, including its examples and their respective modules.
        /// </summary>
        /// <param name="name">Uniquely identifies the feature item to get.</param>
        Task<FeatureItem> GetFeatureItemAsync(string name);
        /// <summary>
        /// Gets metadata for the latest release and prerelease tags.
        /// </summary>
        Task<IEnumerable<Tag>> GetLatestTagsAsync();
    }
}
