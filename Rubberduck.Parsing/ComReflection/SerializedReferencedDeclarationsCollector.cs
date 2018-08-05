using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.ComReflection
{
    public class SerializedReferencedDeclarationsCollector : IReferencedDeclarationsCollector
    {
        private readonly string _serializedDeclarationsPath;

        public SerializedReferencedDeclarationsCollector(string serializedDeclarationsPath = null)
        {
            _serializedDeclarationsPath = serializedDeclarationsPath ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck", "declarations");
        }

        public (IReadOnlyCollection<Declaration> declarations, Dictionary<IList<string>, Declaration> coClasses, SerializableProject serializableProject)
            CollectDeclarations(IReference reference)
        {
            if (!SerializedVersionExists(reference))
            {
                return (new List<Declaration>(), new Dictionary<IList<string>, Declaration>(), null);
            }

            var (declarations, coClasses) = LoadDeclarationsFromXml(reference);
            return (declarations, coClasses, null);
        }

        private bool SerializedVersionExists(IReference reference)
        {
            if (!Directory.Exists(_serializedDeclarationsPath))
            {
                return false;
            }

            //TODO: This is naively based on file name for now - this should attempt to deserialize any SerializableProject.Nodes in the directory and test for equity.
            var testFile = FilePath(reference);
            return File.Exists(testFile);
        }

        private string FilePath(IReference reference)
        {
            return Path.Combine(_serializedDeclarationsPath, FileName(reference));
        }

        private string FileName(IReference reference)
        {
            return $"{reference.Name}.{reference.Major}.{reference.Minor}.xml";
        }

        public (IReadOnlyCollection<Declaration> declarations, Dictionary<IList<string>, Declaration> coClasses) LoadDeclarationsFromXml(IReference reference)
        {
            var file = FilePath(reference);
            var reader = new XmlPersistableDeclarations();
            var deserialized = reader.Load(file);

            var declarations = deserialized.Unwrap();
            var coClasses = new Dictionary<IList<string>, Declaration>();
            foreach (var members in declarations.Where(d => d.DeclarationType != DeclarationType.Project &&
                                                            d.ParentDeclaration.DeclarationType == DeclarationType.ClassModule &&
                                                            ProceduralTypes.Contains(d.DeclarationType))
                .GroupBy(d => d.ParentDeclaration))
            {
                coClasses[members.Select(m => m.IdentifierName).ToList()] = members.Key;
            }
            return (declarations, coClasses);
        }

        private static readonly HashSet<DeclarationType> ProceduralTypes =
            new HashSet<DeclarationType>(new[]
            {
                DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet,
                DeclarationType.PropertyLet, DeclarationType.PropertySet
            });
    }
}
