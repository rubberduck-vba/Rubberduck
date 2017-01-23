using System;
using System.IO;
using System.Runtime.Serialization;
using System.Text;
using System.Xml;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Parsing.ComReflection
{
    public class XmlPersistableDeclarations : IPersistable<SerializableProject>
    {
        public void Persist(string path, SerializableProject tree)
        {
            if (string.IsNullOrEmpty(path)) { throw new InvalidOperationException(); }

            var xmlSettings = new XmlWriterSettings
            {
                NamespaceHandling = NamespaceHandling.OmitDuplicates,
                Encoding = Encoding.UTF8,
                //Indent = true
            };

            using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            using (var xmlWriter = XmlWriter.Create(stream, xmlSettings))
            using (var writer = XmlDictionaryWriter.CreateDictionaryWriter(xmlWriter))
            {
                writer.WriteStartDocument();
                var settings = new DataContractSerializerSettings {RootNamespace = XmlDictionaryString.Empty};
                var serializer = new DataContractSerializer(typeof(SerializableProject), settings);
                serializer.WriteObject(writer, tree);
            }
        }

        public SerializableProject Load(string path)
        {
            if (string.IsNullOrEmpty(path)) { throw new InvalidOperationException(); }
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                return Load(stream);
            }
        }

        public SerializableProject Load(Stream stream)
        {
            if (stream == null) { throw new ArgumentNullException(); }
            using (var xmlReader = XmlReader.Create(stream))
            using (var reader = XmlDictionaryReader.CreateDictionaryReader(xmlReader))
            {
                var serializer = new DataContractSerializer(typeof(SerializableProject));
                return (SerializableProject)serializer.ReadObject(reader);
            }
        }
    }
}