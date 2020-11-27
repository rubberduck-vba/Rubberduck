using System;
using System.Diagnostics.CodeAnalysis;
using MemoryStream = System.IO.MemoryStream;
using Path = System.IO.Path;
using StreamWriter = System.IO.StreamWriter;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.IO.Abstractions;

namespace Rubberduck.SettingsProvider
{
    internal class XmlPersistenceService<T> : XmlPersistenceServiceBase<T> 
        where T : class, IEquatable<T>, new()
    {
        private const string DefaultConfigFile = "rubberduck.config";

        public XmlPersistenceService(
            IPersistencePathProvider pathProvider,
            IFileSystem fileSystem) 
            : base(pathProvider, fileSystem) { }

        protected override string FilePath => Path.Combine(RootPath, DefaultConfigFile);

        protected override T Read(string path)
        {
            var doc = GetConfigurationDoc(path);
            var node = GetNodeByName(doc, typeof(T).Name);
            if (node == null)
            {
                return default;
            }

            using (var reader = node.CreateReader())
            {
                var deserializer = new XmlSerializer(typeof(T));
                try
                {
                    return (T)deserializer.Deserialize(reader);
                }
                catch
                {
                    return default;
                }
            }
        }

        //This is fine. StreamWriter disposes the MemoryStream, but calling twice is a NOP.
        [SuppressMessage("Microsoft.Usage", "CA2202:Do not dispose objects multiple times")] 
        protected override void Write(T toSerialize, string path)
        {
            var doc = GetConfigurationDoc(path);
            var node = GetNodeByName(doc, typeof(T).Name);
            using (var stream = new MemoryStream())
            using (var writer = new StreamWriter(stream))
            {
                var serializer = new XmlSerializer(typeof(T));
                serializer.Serialize(writer, toSerialize, EmptyNamespace);
                var settings = XElement.Parse(OutputEncoding.GetString(stream.ToArray()), LoadOptions.SetBaseUri);

                if (node != null)
                {
                    node.ReplaceWith(settings);
                }
                else
                {
                    GetNodeByName(doc, RootElement).Add(settings);
                }
            }

            using (var xml = XmlWriter.Create(path, OutputXmlSettings))
            {
                doc.WriteTo(xml);
            }
        }
    }
}
