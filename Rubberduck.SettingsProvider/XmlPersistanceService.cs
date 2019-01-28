using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace Rubberduck.SettingsProvider
{
    public class XmlPersistanceService<T> : XmlPersistanceServiceBase<T> where T : class, IEquatable<T>, new()
    {
        public override T Load(T toDeserialize)
        {
            var defaultOutput = CachedOrNotFound();
            if (defaultOutput != null)
            {
                return defaultOutput;
            }

            var doc = GetConfigurationDoc(FilePath);
            var node = GetNodeByName(doc, typeof(T).Name);
            if (node == null)
            {
                return FailedLoadReturnValue();
            }

            using (var reader = node.CreateReader())
            {
                var deserializer = new XmlSerializer(typeof(T));
                try
                {
                    Cached = (T)Convert.ChangeType(deserializer.Deserialize(reader), typeof(T));
                    return Cached;
                }
                catch
                {
                    return FailedLoadReturnValue();
                }
            }
        }

        [SuppressMessage("Microsoft.Usage", "CA2202:Do not dispose objects multiple times")] //This is fine. StreamWriter disposes the MemoryStream, but calling twice is a NOP.
        public override void Save(T toSerialize)
        {
            var doc = GetConfigurationDoc(FilePath);
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

            EnsurePathExists();

            using (var xml = XmlWriter.Create(FilePath, OutputXmlSettings))
            {
                doc.WriteTo(xml);
                Cached = toSerialize;
            }
        }
    }
}
