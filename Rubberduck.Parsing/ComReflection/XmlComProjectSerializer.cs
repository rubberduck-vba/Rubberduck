using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using FileAccess = System.IO.FileAccess;
using FileMode = System.IO.FileMode;
using Stream = System.IO.Stream;
using System.Runtime.Serialization;
using System.Xml;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor;
using System.IO.Abstractions;

namespace Rubberduck.Parsing.ComReflection
{
    public class XmlComProjectSerializer : IComProjectSerializationProvider
    {
        public readonly string DefaultSerializationPath;
        private readonly IFileSystem _fileSystem;

        private XmlComProjectSerializer(IFileSystem fileSystem, string defaultPath, string path)
        {
            _fileSystem = fileSystem;
            DefaultSerializationPath = defaultPath;
            Target = path ?? DefaultSerializationPath;
        }

        public XmlComProjectSerializer(
            IPersistencePathProvider pathProvider,
            IFileSystem fileSystem)
            : this(fileSystem, pathProvider.DataFolderPath("Declarations"), null)
        { }

        public XmlComProjectSerializer(IFileSystem fileSystem, string path)
            : this(fileSystem, path, path)
        { }

        private static readonly XmlReaderSettings ReaderSettings = new XmlReaderSettings
        {
            CheckCharacters = false
        };

        private static readonly XmlWriterSettings WriterSettings = CreateWriterSettings();
        
        private static XmlWriterSettings CreateWriterSettings()
        {
            var settings = new XmlWriterSettings
            {
                NamespaceHandling = NamespaceHandling.OmitDuplicates,
                CheckCharacters = false
            };

            PrettifyWriter(ref settings);

            return settings;
        }

        [Conditional("PRETTY_XML")]
        private static void PrettifyWriter(ref XmlWriterSettings settings)
        {
            settings.Indent = true;
            settings.IndentChars = ("\t");
            settings.NewLineChars = Environment.NewLine;
        }

        public string Target { get; }

        public bool SerializedVersionExists(ReferenceInfo reference)
        {
            if (!_fileSystem.Directory.Exists(Target))
            {
                return false;
            }

            //TODO: This is naively based on file name for now - this should attempt to deserialize any SerializableProject.Nodes in the directory and test for equity.
            var testFile = _fileSystem.Path.Combine(Target, FileName(reference));
            return _fileSystem.File.Exists(testFile);
        }

        [SuppressMessage("Microsoft.Usage", "CA2202:Do not dispose objects multiple times")] //This is fine. XmlWriter disposes the FileStream, but calling twice is a NOP.
        public void SerializeProject(ComProject project)
        {
            var filepath = _fileSystem.Path.Combine(Target, FileName(project));

            using (var stream = _fileSystem.FileStream.Create(filepath, FileMode.Create, FileAccess.Write))
            using (var xmlWriter = XmlWriter.Create(stream, WriterSettings))
            using (var writer = XmlDictionaryWriter.CreateDictionaryWriter(xmlWriter))
            {
                writer.WriteStartDocument();
                var settings = new DataContractSerializerSettings
                {
                    RootNamespace = XmlDictionaryString.Empty,
                    PreserveObjectReferences = true
                };
                var serializer = new DataContractSerializer(typeof(ComProject), settings);
                serializer.WriteObject(writer, project);
            }
        }

        public ComProject DeserializeProject(ReferenceInfo reference)
        {
            var filepath = _fileSystem.Path.Combine(Target, FileName(reference));

            if (string.IsNullOrEmpty(filepath))
            {
                throw new InvalidOperationException();
            }

            using (var stream = _fileSystem.FileStream.Create(filepath, FileMode.Open, FileAccess.Read))
            {
                return Load(stream);
            }
        }

        private static ComProject Load(Stream stream)
        {
            if (stream == null)
            {
                throw new ArgumentNullException();
            }

            using (var xmlReader = XmlReader.Create(stream, ReaderSettings))
            using (var reader = XmlDictionaryReader.CreateDictionaryReader(xmlReader))
            {
                var serializer = new DataContractSerializer(typeof(ComProject));
                return (ComProject)serializer.ReadObject(reader);
            }
        }

        private static string FileName(ReferenceInfo reference)
        {
            return $"{reference.Name}.{reference.Major}.{reference.Minor}.xml";
        }

        private static string FileName(ComProject project)
        {
            return $"{project.Name}.{project.MajorVersion}.{project.MinorVersion}.xml";
        }
    }
}
