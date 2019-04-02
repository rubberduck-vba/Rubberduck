using System;

namespace Rubberduck.Parsing.Common
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExperimentalAttribute : Attribute
    {
        public ExperimentalAttribute(string resource)
        {
            Resource = resource;
        }

        /// <summary>
        /// Resource key to look up in <see cref="Rubberduck.Resources.Experimental.ExperimentalNames"/>.
        /// Also serves as a unique identifier to distinguish experimental features from one another.
        /// </summary>
        public string Resource { get; }
    }
}