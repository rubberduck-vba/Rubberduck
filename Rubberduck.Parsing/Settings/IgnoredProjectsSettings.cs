using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;

namespace Rubberduck.Parsing.Settings
{
    public interface IIgnoredProjectsSettings
    {
        List<string> IgnoredProjectPaths { get; set; }
    }

    [DataContract]
    public class IgnoredProjectsSettings : IIgnoredProjectsSettings, IEquatable<IgnoredProjectsSettings>
    {
        [DataMember(IsRequired = true)]
        public List<string> IgnoredProjectPaths { get; set; } = new List<string>();
        
        public bool Equals(IgnoredProjectsSettings other)
        {
            if (ReferenceEquals(this, other))
            {
                return true;
            }

            if (other?.IgnoredProjectPaths is null)
            {
                return false;
            }

            return other.IgnoredProjectPaths.Count == IgnoredProjectPaths.Count
                   && other.IgnoredProjectPaths.All(IgnoredProjectPaths.Contains);
        }
    }
}