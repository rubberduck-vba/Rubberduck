using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.ComReflection
{
    public class SerializedReferencedDeclarationsCollector : ReferencedDeclarationsCollectorBase
    {
        private readonly IComProjectDeserializer _deserializer;

        public SerializedReferencedDeclarationsCollector(IDeclarationsFromComProjectLoader declarationsFromComProjectLoader, IComProjectDeserializer deserializer)
        :base(declarationsFromComProjectLoader)
        {
            _deserializer = deserializer;
        }

        public override IReadOnlyCollection<Declaration> CollectedDeclarations(ReferenceInfo reference)
        {
            if (!_deserializer.SerializedVersionExists(reference))
            {
                return new List<Declaration>();
            }

            return LoadDeclarationsFromProvider(reference);
        }

        private IReadOnlyCollection<Declaration> LoadDeclarationsFromProvider(ReferenceInfo reference)
        {
            var type = _deserializer.DeserializeProject(reference);
            return LoadDeclarationsFromComProject(type); 
        }
    }
}
