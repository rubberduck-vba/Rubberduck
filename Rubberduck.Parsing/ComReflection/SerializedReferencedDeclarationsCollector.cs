using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.ComReflection
{
    public class SerializedReferencedDeclarationsCollector : ReferencedDeclarationsCollectorBase
    {
        private readonly IComProjectSerializationProvider _serializer;

        public SerializedReferencedDeclarationsCollector(string serializedDeclarationsPath = null)
        {
            _serializer = new XmlComProjectSerializer(serializedDeclarationsPath);
        }

        public override IReadOnlyCollection<Declaration> CollectedDeclarations(ReferenceInfo reference)
        {
            if (!_serializer.SerializedVersionExists(reference))
            {
                return new List<Declaration>();
            }

            return LoadDeclarationsFromProvider(reference);
        }

        private IReadOnlyCollection<Declaration> LoadDeclarationsFromProvider(ReferenceInfo reference)
        {
            var type = _serializer.DeserializeProject(reference);
            return LoadDeclarationsFromComProject(type); 
        }
    }
}
