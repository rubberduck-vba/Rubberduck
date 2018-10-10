using System;
using System.Collections.Generic;
using System.IO;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.ComReflection
{
    public class SerializedReferencedDeclarationsCollector : IReferencedDeclarationsCollector
    {
        private readonly string _serializedDeclarationsPath;

        public SerializedReferencedDeclarationsCollector(string serializedDeclarationsPath = null)
        {
            _serializedDeclarationsPath = serializedDeclarationsPath ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck", "declarations");
        }

        public IReadOnlyCollection<Declaration> CollectedDeclarations(ReferenceInfo reference)
        {
            if (!SerializedVersionExists(reference))
            {
                return new List<Declaration>();
            }

            return LoadDeclarationsFromXml(reference);
        }

        private bool SerializedVersionExists(ReferenceInfo reference)
        {
            if (!Directory.Exists(_serializedDeclarationsPath))
            {
                return false;
            }

            //TODO: This is naively based on file name for now - this should attempt to deserialize any SerializableProject.Nodes in the directory and test for equity.
            var testFile = FilePath(reference);
            return File.Exists(testFile);
        }

        private string FilePath(ReferenceInfo reference)
        {
            return Path.Combine(_serializedDeclarationsPath, FileName(reference));
        }

        private string FileName(ReferenceInfo reference)
        {
            return $"{reference.Name}.{reference.Major}.{reference.Minor}.xml";
        }

        public IReadOnlyCollection<Declaration> LoadDeclarationsFromXml(ReferenceInfo reference)
        {
            var file = FilePath(reference);
            var reader = new XmlPersistableDeclarations();
            var deserialized = reader.Load(file);
            return deserialized.Unwrap();
        }

        private static readonly HashSet<DeclarationType> ProceduralTypes =
            new HashSet<DeclarationType>(new[]
            {
                DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet,
                DeclarationType.PropertyLet, DeclarationType.PropertySet
            });
    }
}
