using System.Collections.Generic;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.VBEditor;

namespace RubberduckTests.Mocks
{
    public class StaticCachingComDeserializerDecorator : IComProjectDeserializer
    {
        private static readonly Dictionary<ReferenceInfo, ComProject> CachedReferences = new Dictionary<ReferenceInfo, ComProject>();

        private static readonly object _lockObject = new object();

        private IComProjectDeserializer _baseDeserializer;

        public StaticCachingComDeserializerDecorator(IComProjectDeserializer baseDeserializer)
        {
            _baseDeserializer = baseDeserializer;
        }

        public bool SerializedVersionExists(ReferenceInfo reference)
        {
            lock(_lockObject)
            {
                return CachedReferences.TryGetValue(reference, out _) || _baseDeserializer.SerializedVersionExists(reference);
            }
        }

        public ComProject DeserializeProject(ReferenceInfo reference)
        {
            lock (_lockObject)
            {
                if (CachedReferences.TryGetValue(reference, out var cachedProject))
                {
                    return cachedProject;
                }

                var comProject = _baseDeserializer.DeserializeProject(reference);
                CachedReferences[reference] = comProject;
                return comProject;
            }
        } 
    }
}