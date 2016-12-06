using System;
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

        private static readonly XmlSerializerNamespaces EmptyNamespace =
            new XmlSerializerNamespaces(new[] { XmlQualifiedName.Empty });
        
        private static readonly XmlWriterSettings OutputXmlSettings = new XmlWriterSettings
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
            var type = typeof(T);

            if (!File.Exists(FilePath))
            {
                return FailedLoadReturnValue();
            }
            var doc = GetConfigurationDoc(FilePath);
            
            var node = doc.Descendants().FirstOrDefault(e => e.Name.LocalName.Equals(type.Name));
            if (node == null)
            {
                return FailedLoadReturnValue();
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
                    return FailedLoadReturnValue();
                }
            }  
        }

            private static T FailedLoadReturnValue()
            {
                return (T)Convert.ChangeType(null, typeof(T));
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
                serializer.Serialize(writer, toSerialize, EmptyNamespace);
                var settings = XElement.Parse(_outputEncoding.GetString(stream.ToArray()), LoadOptions.SetBaseUri);
                if (node != null)
                {
                    node.ReplaceWith(settings);
                }
                else
                {
                    var parent = doc.Descendants().FirstOrDefault(e => e.Name.LocalName.Equals(RootElement));
                    // ReSharper disable once PossibleNullReferenceException
                    parent.Add(settings);
                }                
            }

            if (!Directory.Exists(_rootPath))
            {
                Directory.CreateDirectory(_rootPath);
            }

            using (var xml = XmlWriter.Create(FilePath, OutputXmlSettings))
            {
                doc.WriteTo(xml);
            }
        }

        private static XDocument GetConfigurationDoc(string file)
        {
            XDocument output;
            try
            {
                output = XDocument.Load(file);
                if (output.Descendants().FirstOrDefault(e => e.Name.LocalName.Equals(RootElement)) != null)
                {
                    return output;
                }
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch { }
            
            output = new XDocument();
            output.Add(new XElement(RootElement));
            return output;
        }
    }
}
