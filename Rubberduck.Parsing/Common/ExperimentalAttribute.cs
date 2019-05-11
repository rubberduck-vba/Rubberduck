using System;

namespace Rubberduck.Parsing.Common
{
    /// <summary>
    /// Marks a class as belonging to an experimental feature.
    /// The feature is identified by the resource key in <see cref="Resources.Experimentals.ExperimentalNames"/> that describes it to the user.
    /// Features marked as experimental are excluded from IoC configuration unless the user has explicitly enabled them.
    /// <para>
    /// See also: <seealso cref="DisabledAttribute"/>
    /// </para>
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class ExperimentalAttribute : Attribute
    {
        public ExperimentalAttribute(string resource)
        {
            Resource = resource;
        }

        /// <summary>
        /// Resource key to look up in <see cref="Resources.Experimentals.ExperimentalNames"/>.
        /// Also serves as a unique identifier to distinguish experimental features from one another.
        /// </summary>
        public string Resource { get; }
    }
}