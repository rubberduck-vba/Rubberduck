using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace Rubberduck.SettingsProvider
{
    internal class XmlPersistenceService<T> : XmlPersistenceServiceBase<T> 
        where T : class, IEquatable<T>, new()
    {
        public XmlPersistenceService(IPersistencePathProvider pathProvider) : base(pathProvider) { }

        public override T Load(T toDeserialize, string nonDefaultFilePath = null)
        {
            var filePath = string.IsNullOrWhiteSpace(nonDefaultFilePath) ? FilePath : nonDefaultFilePath;
            var doc = GetConfigurationDoc(filePath);
            var node = GetNodeByName(doc, typeof(T).Name);
            if (node == null)
            {
                return Cached;
            }

            using (var reader = node.CreateReader())
            {
                var deserializer = new XmlSerializer(typeof(T));
                try
                {
                    Cached = (T)deserializer.Deserialize(reader);
                    return Cached;
                }
                catch
                {
                    return FailedLoadReturnValue();
                }
            }
        }

        [SuppressMessage("Microsoft.Usage", "CA2202:Do not dispose objects multiple times")] //This is fine. StreamWriter disposes the MemoryStream, but calling twice is a NOP.
        public override void Save(T toSerialize, string nonDefaultFilePath = null)
        {
            var filePath = string.IsNullOrWhiteSpace(nonDefaultFilePath) ? FilePath : nonDefaultFilePath;
            var doc = GetConfigurationDoc(filePath);
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

            EnsurePathExists(filePath);

            using (var xml = XmlWriter.Create(filePath, OutputXmlSettings))
            {
                doc.WriteTo(xml);
                Cached = toSerialize;
            }
        }
    }
}
