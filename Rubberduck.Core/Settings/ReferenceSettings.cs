using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Xml.Serialization;
using Rubberduck.VBEditor;

namespace Rubberduck.Settings
{
    public interface IReferenceSettings
    {
        int RecentReferencesTracked { get; set; }
        List<ReferenceInfo> GetRecentReferencesForHost(string host);
        void UpdateRecentReferencesForHost(string host, List<ReferenceInfo> references);
        List<ReferenceInfo> GetPinnedReferencesForHost(string host);
        void UpdatePinnedReferencesForHost(string host, List<ReferenceInfo> references);
        void PinReference(ReferenceInfo reference, string host = null);
        void TrackUsage(ReferenceInfo reference, string host = null);
    }

    [DataContract]
    [KnownType(typeof(ReferenceInfo))]
    [KnownType(typeof(ReferenceUsage))]
    public class ReferenceSettings : IReferenceSettings, IEquatable<ReferenceSettings>
    {
        [DataMember(IsRequired = true)]
        [XmlElement(ElementName = "RecentReferences")]
        private List<HostUsages> _recent;

        [DataMember(IsRequired = true)]
        [XmlElement(ElementName = "PinnedProjects")]
        private List<HostPins> _pinned;

        [OnDeserialized]
        private void DeserializationLoad(StreamingContext context)
        {
            RecentProjectReferences = _recent.ToDictionary(usage => usage.Host, usage => usage.Usages);
            PinnedProjectReferences = _pinned.ToDictionary(usage => usage.Host, usage => usage.Pins);
        }

        [OnSerializing]
        private void SerializationPrep(StreamingContext context)
        {
            _recent = new List<HostUsages>(RecentProjectReferences.Select(recent => new HostUsages(recent.Key, recent.Value)));
            _pinned = new List<HostPins>(PinnedProjectReferences.Select(recent => new HostPins(recent.Key, recent.Value)));
        }

        [DataMember(IsRequired = true)]
        public int RecentReferencesTracked { get; set; }

        [DataMember(IsRequired = true)]
        protected List<ReferenceUsage> RecentLibraryReferences { get; private set; } = new List<ReferenceUsage>();

        [DataMember(IsRequired = true)]
        protected List<ReferenceInfo> PinnedLibraryReferences { get; private set; } = new List<ReferenceInfo>();
     
        protected Dictionary<string, List<ReferenceInfo>> PinnedProjectReferences { get; private set; } = new Dictionary<string, List<ReferenceInfo>>();
        protected Dictionary<string, List<ReferenceUsage>> RecentProjectReferences { get; private set; } = new Dictionary<string, List<ReferenceUsage>>();

        public void PinReference(ReferenceInfo reference, string host = null)
        {
            if (string.IsNullOrEmpty(host))
            {
                PinnedLibraryReferences.Add(reference);
                return;
            }

            var key = host.ToUpperInvariant();
            if (!PinnedProjectReferences.ContainsKey(key))
            {
                PinnedProjectReferences.Add(key, new List<ReferenceInfo> { reference });
            }
            else
            {
                PinnedProjectReferences[key].Add(reference);
            }
        }

        public void TrackUsage(ReferenceInfo reference, string host = null)
        {
            var use = new ReferenceUsage(reference);
            if (string.IsNullOrEmpty(host))
            {
                RecentLibraryReferences.Add(use);
                RecentLibraryReferences = RecentLibraryReferences
                    .OrderByDescending(usage => usage.Timestamp)
                    .Take(RecentReferencesTracked).ToList();
                return;
            }

            var key = host.ToUpperInvariant();
            if (!RecentProjectReferences.ContainsKey(key))
            {
                RecentProjectReferences.Add(key, new List<ReferenceUsage> { use });
            }
            else
            {
                var recent = RecentProjectReferences[key];
                recent.RemoveAll(usage => usage.Matches(reference));
                recent.Add(use);
                RecentProjectReferences[key] = recent;
            }

            RecentProjectReferences[key] = RecentProjectReferences[key]
                .OrderByDescending(usage => usage.Timestamp)
                .Take(RecentReferencesTracked).ToList();
        }

        public bool Equals(ReferenceSettings other)
        {
            if (other is null || 
                RecentReferencesTracked != other.RecentReferencesTracked || 
                !PinnedLibraryReferences.OrderBy(_ => _).SequenceEqual(other.PinnedLibraryReferences.OrderBy(_ => _)) ||
                !RecentLibraryReferences.OrderBy(_ => _).SequenceEqual(other.RecentLibraryReferences.OrderBy(_ => _)))
            {
                return false;
            }

            foreach (var host in PinnedProjectReferences)
            {
                if (!other.PinnedProjectReferences.ContainsKey(host.Key) ||
                    !host.Value.OrderBy(_ => _).SequenceEqual(other.PinnedProjectReferences[host.Key].OrderBy(_ => _)))
                {
                    return false;
                }
            }

            foreach (var host in RecentProjectReferences)
            {
                if (!other.RecentProjectReferences.ContainsKey(host.Key) ||
                    !host.Value.OrderBy(usage => usage.Timestamp).Select(usage => usage.Reference)
                        .SequenceEqual(other.RecentProjectReferences[host.Key].OrderBy(usage => usage.Timestamp).Select(usage => usage.Reference)))
                {
                    return false;
                }
            }

            return true;
        }

        public List<ReferenceInfo> GetPinnedReferencesForHost(string host)
        {
            var key = host.ToUpperInvariant();
            return PinnedLibraryReferences.Union(PinnedProjectReferences.ContainsKey(key)
                ? PinnedProjectReferences[key].ToList()
                : new List<ReferenceInfo>()).ToList();
        }

        public List<ReferenceInfo> GetRecentReferencesForHost(string host)
        {
            var key = host.ToUpperInvariant();
            return RecentLibraryReferences
                .Concat(RecentProjectReferences.ContainsKey(key)
                    ? RecentProjectReferences[key]
                    : new List<ReferenceUsage>()).OrderBy(reference => reference.Timestamp)
                .Select(reference => reference.Reference)
                .Take(RecentReferencesTracked).ToList();
        }

        public void UpdatePinnedReferencesForHost(string host, List<ReferenceInfo> pinned)
        {
            foreach (var reference in pinned)
            {
                PinReference(reference, reference.Guid.Equals(Guid.Empty) ? host : string.Empty);
            }
        }

        public void UpdateRecentReferencesForHost(string host, List<ReferenceInfo> references)
        {
            foreach (var usage in references)
            {
                TrackUsage(usage, usage.Guid.Equals(Guid.Empty) ? host : string.Empty);
            }
        }

        [DataContract]
        [KnownType(typeof(ReferenceInfo))]
        protected class ReferenceUsage
        {
            [DataMember(IsRequired = true)]
            public ReferenceInfo Reference { get; protected set; }

            [DataMember(IsRequired = true)]
            public DateTime Timestamp { get; protected set; } = DateTime.Now;

            public ReferenceUsage(ReferenceInfo reference)
            {
                Reference = reference;
            }

            public bool Matches(ReferenceInfo other)
            {
                return Reference.FullPath.Equals(other.FullPath, StringComparison.OrdinalIgnoreCase) ||
                       Reference.Guid.Equals(other.Guid) &&
                       Reference.Major == other.Major &&
                       Reference.Minor == other.Minor;
            }
        }

        [DataContract]
        private struct HostUsages
        {
            public string Host { get; }
            public List<ReferenceUsage> Usages { get; }

            public HostUsages(string host, List<ReferenceUsage> usages)
            {
                Host = host;
                Usages = usages;
            }
        }

        [DataContract]
        private struct HostPins
        {
            public string Host { get; }
            public List<ReferenceInfo> Pins { get; }

            public HostPins(string host, List<ReferenceInfo> usages)
            {
                Host = host;
                Pins = usages;
            }
        }
    }
}
