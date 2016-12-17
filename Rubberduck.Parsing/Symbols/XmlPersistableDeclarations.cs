using System;
using System.IO;
using System.Runtime.Serialization;
using System.Text;
using System.Xml;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Parsing.Symbols
{
    public class XmlPersistableDeclarations : IPersistable<SerializableDeclarationTree>
    {
        public void Persist(string path, SerializableDeclarationTree tree)
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
                var serializer = new DataContractSerializer(typeof (SerializableDeclarationTree), settings);
                serializer.WriteObject(writer, tree);
            }
        }

        public SerializableDeclarationTree Load(string path)
        {
            if (string.IsNullOrEmpty(path)) { throw new InvalidOperationException(); }
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read))
            using (var xmlReader = XmlReader.Create(stream))
            using (var reader = XmlDictionaryReader.CreateDictionaryReader(xmlReader))
            {
                var serializer = new DataContractSerializer(typeof(SerializableDeclarationTree));
                return (SerializableDeclarationTree)serializer.ReadObject(reader);
            }
        }
    }
}