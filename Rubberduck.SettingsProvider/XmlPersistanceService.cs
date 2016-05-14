using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace Rubberduck.SettingsProvider
{
    public class XmlPersistanceService<T> : IFilePersistanceService<T> where T : new()
    {
        private readonly string _rootPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck");
        private readonly UTF8Encoding _outputEncoding = new UTF8Encoding(false);
        private const string DefaultConfigFile = "rubberduck.config";
        private const string RootElement = "Configuration";

        private readonly XmlSerializerNamespaces _emptyNamespace =
            new XmlSerializerNamespaces(new[] { new XmlQualifiedName(string.Empty, string.Empty) });
        
        private readonly XmlWriterSettings _outputXmlSettings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            Indent = true,
        };

        private string _filePath;
        public string FilePath
        {
            get { return _filePath ?? Path.Combine(_rootPath, DefaultConfigFile); }
            set { _filePath = value; }
        }

        public T Load(T toDeserialize)
        {
            var doc = GetConfigurationDoc(FilePath);
            var type = typeof(T);
            var node = doc.Descendants().FirstOrDefault(e => e.Name.LocalName.Equals(type.Name));
            if (node == null)
            {
                return (T)Convert.ChangeType(null, type);
            }

            using (var reader = node.CreateReader())
            {
                var deserializer = new XmlSerializer(type);
                try
                {
                    var output = deserializer.Deserialize(reader);
                    return (T)Convert.ChangeType(output, type);
                }
                catch
                {
                    return (T)Convert.ChangeType(null, type);
                }
            }  
        }

        public void Save(T toSerialize)
        {
            var doc = GetConfigurationDoc(FilePath);
            var type = typeof(T);
            var node = doc.Descendants().FirstOrDefault(e => e.Name.LocalName.Equals(type.Name));
            
            using (var stream = new MemoryStream())
            using (var writer = new StreamWriter(stream))
            {
                var serializer = new XmlSerializer(type);
                serializer.Serialize(writer, toSerialize, _emptyNamespace);
                var settings = XElement.Parse(_outputEncoding.GetString(stream.ToArray()), LoadOptions.SetBaseUri);
                if (node != null)
                {
                    node.ReplaceWith(settings);
                }
                else
                {
                    var parent = doc.Descendants().FirstOrDefault(e => e.Name.LocalName.Equals(RootElement));
                    Debug.Assert(parent != null);
                    parent.Add(settings);
                }                
            }
            
            using (var xml = XmlWriter.Create(FilePath, _outputXmlSettings))
            {
                doc.WriteTo(xml);
            }
        }

        private static XDocument GetConfigurationDoc(string file)
        {
            try
            {
                return XDocument.Load(file);
            }
            catch
            {
                var output = new XDocument();
                var root = output.Descendants(RootElement).FirstOrDefault();
                if (root == null)
                {
                    output.Add(new XElement(RootElement));
                    root = output.Descendants(RootElement).FirstOrDefault();
                }
                Debug.Assert(root != null);
                return output;
            }
        }
    }
}