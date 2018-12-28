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
        bool FixBrokenReferences { get; set; }
        bool AddToRecentOnReferenceEvents { get; set; }
        List<string> ProjectPaths { get; set; }
        List<ReferenceInfo> GetRecentReferencesForHost(string host);
        void UpdateRecentReferencesForHost(string host, List<ReferenceInfo> references);
        List<ReferenceInfo> GetPinnedReferencesForHost(string host);
        void UpdatePinnedReferencesForHost(string host, List<ReferenceInfo> references);
        void PinReference(ReferenceInfo reference, string host = null);
        void TrackUsage(ReferenceInfo reference, string host = null);
        bool IsPinnedProject(string filePath, string host);
        bool IsRecentProject(string filePath, string host);
    }

    [DataContract]
    [KnownType(typeof(ReferenceInfo))]
    [KnownType(typeof(ReferenceUsage))]
    public class ReferenceSettings : IReferenceSettings, IEquatable<ReferenceSettings>
    {
        public const int RecentTrackingLimit = 50;

        [DataMember(IsRequired = true)]
        [XmlElement(ElementName = "RecentReferences")]
        private List<HostUsages> _recent;

        [DataMember(IsRequired = true)]
        [XmlElement(ElementName = "PinnedProjects")]
        private List<HostPins> _pinned;

        [OnDeserialized]
        private void DeserializationLoad(StreamingContext context)
        {
            RecentProjectReferences = _recent.Take(RecentTrackingLimit).ToDictionary(usage => usage.Host, usage => usage.Usages);
            PinnedProjectReferences = _pinned.ToDictionary(usage => usage.Host, usage => usage.Pins);
        }

        [OnSerializing]
        private void SerializationPrep(StreamingContext context)
        {
            _recent = new List<HostUsages>(RecentProjectReferences.Select(recent => new HostUsages(recent.Key, recent.Value)).Take(RecentTrackingLimit));
            _pinned = new List<HostPins>(PinnedProjectReferences.Select(recent => new HostPins(recent.Key, recent.Value)));
        }

        public ReferenceSettings() { }

        public ReferenceSettings(ReferenceSettings other)
        {
            RecentReferencesTracked = other.RecentReferencesTracked;
            FixBrokenReferences = other.FixBrokenReferences;
            AddToRecentOnReferenceEvents = other.AddToRecentOnReferenceEvents;
            ProjectPaths = new List<string>(other.ProjectPaths);
            other.SerializationPrep(new StreamingContext(StreamingContextStates.All));
            _recent = other._recent.Select(use => new HostUsages(use)).ToList();
            RecentLibraryReferences = other.RecentLibraryReferences.ToList();
            _pinned = other._pinned.Select(pin => new HostPins(pin)).ToList();
            PinnedLibraryReferences = other.PinnedLibraryReferences.ToList();
            DeserializationLoad(new StreamingContext(StreamingContextStates.All));
        }

        private int _tracked;

        [DataMember(IsRequired = true)]
        public int RecentReferencesTracked
        {
            get => _tracked;
            set => _tracked = value < 0 ? 0 : Math.Min(value, RecentTrackingLimit);
        }

        [DataMember(IsRequired = true)]
        public bool FixBrokenReferences { get; set; }

        [DataMember(IsRequired = true)]
        public bool AddToRecentOnReferenceEvents { get; set; }

        [DataMember(IsRequired = true)]
        public List<string> ProjectPaths { get; set; } = new List<string>();

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

        public bool IsPinnedProject(string filePath, string host)
        {
            var key = host.ToUpperInvariant();
            return PinnedProjectReferences.ContainsKey(key) && 
                   PinnedProjectReferences[key]
                       .Select(pin => pin.FullPath)
                       .Contains(filePath, StringComparer.OrdinalIgnoreCase);
        }

        public bool IsRecentProject(string filePath, string host)
        {
            var key = host.ToUpperInvariant();
            return RecentProjectReferences.ContainsKey(key) &&
                   RecentProjectReferences[key]
                       .Select(usage => usage.Reference.FullPath)
                       .Contains(filePath, StringComparer.OrdinalIgnoreCase);
        }

        public void TrackUsage(ReferenceInfo reference, string host = null)
        {
            var use = new ReferenceUsage(reference);
            if (string.IsNullOrEmpty(host))
            { 
                RecentLibraryReferences.RemoveAll(usage => usage.Matches(reference));
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

        // This is so close to damned near impossible that I was tempted to hard code it false, but it's useful for testing.
        public bool Equals(ReferenceSettings other)
        {
            if (ReferenceEquals(this, other))
            {
                return true;
            }

            if (other is null || 
                RecentReferencesTracked != other.RecentReferencesTracked ||
                PinnedLibraryReferences.Count != other.PinnedLibraryReferences.Count ||
                RecentLibraryReferences.Count != other.RecentLibraryReferences.Count ||
                PinnedLibraryReferences.Any(pin => !other.PinnedLibraryReferences.Any(lib => lib.Equals(pin))) ||
                RecentLibraryReferences.Any(recent => !other.RecentLibraryReferences.Any(lib => lib.Equals(recent))))
            {
                return false;
            }

            foreach (var host in PinnedProjectReferences)
            {
                if (!other.PinnedProjectReferences.ContainsKey(host.Key) ||
                    !(other.PinnedProjectReferences[host.Key] is List<ReferenceInfo> otherHost) ||
                    otherHost.Count != host.Value.Count ||
                    host.Value.Any(pin => !otherHost.Any(lib => lib.Equals(pin))))
                {
                    return false;
                }
            }

            foreach (var host in RecentProjectReferences)
            {
                if (!other.RecentProjectReferences.ContainsKey(host.Key) ||
                    !(other.RecentProjectReferences[host.Key] is List<ReferenceUsage> otherHost) ||
                    otherHost.Count != host.Value.Count ||
                    host.Value.Any(pin => !otherHost.Any(lib => lib.Reference.Equals(pin.Reference) && lib.Timestamp.Equals(pin.Timestamp))))
                {
                    return false;
                }
            }

            return true;
        }

        public List<ReferenceInfo> GetPinnedReferencesForHost(string host)
        {
            var key = host?.ToUpperInvariant() ?? string.Empty;
            return PinnedLibraryReferences.Union(PinnedProjectReferences.ContainsKey(key)
                ? PinnedProjectReferences[key].ToList()
                : new List<ReferenceInfo>()).ToList();
        }

        public List<ReferenceInfo> GetRecentReferencesForHost(string host)
        {
            var key = host?.ToUpperInvariant() ?? string.Empty;
            return RecentLibraryReferences
                .Concat(RecentProjectReferences.ContainsKey(key)
                    ? RecentProjectReferences[key]
                    : new List<ReferenceUsage>()).OrderBy(reference => reference.Timestamp)
                .Select(reference => reference.Reference)
                .Take(RecentReferencesTracked).ToList();
        }

        public void UpdatePinnedReferencesForHost(string host, List<ReferenceInfo> pinned)
        {
            var key = host?.ToUpperInvariant() ?? string.Empty;

            PinnedLibraryReferences.Clear();
            if (PinnedProjectReferences.ContainsKey(key))
            {
                PinnedProjectReferences.Remove(key);
            }
            
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
            public DateTime Timestamp { get; protected set; }

            public ReferenceUsage(ReferenceInfo reference)
            {
                Reference = reference;
                Timestamp = DateTime.Now;
            }

            public bool Matches(ReferenceInfo other)
            {
                return Reference.FullPath.Equals(other.FullPath, StringComparison.OrdinalIgnoreCase) ||
                       !Reference.Guid.Equals(Guid.Empty) &&
                       !other.Guid.Equals(Guid.Empty) &&
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

            public HostUsages(HostUsages other)
            {
                Host = other.Host;
                Usages = other.Usages.ToList();
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

            public HostPins(HostPins other)
            {
                Host = other.Host;
                Pins = other.Pins.ToList();
            }
        }
    }
}
