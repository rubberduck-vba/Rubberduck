using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

// ReSharper disable StaticMemberInGenericType
namespace Rubberduck.SettingsProvider
{
    public abstract class XmlPersistanceServiceBase<T> : IFilePersistanceService<T> where T : class, IEquatable<T>, new()
    {
        private const string DefaultConfigFile = "rubberduck.config";

        protected readonly string RootPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck");
        protected static readonly UTF8Encoding OutputEncoding = new UTF8Encoding(false);        
        protected const string RootElement = "Configuration";

        protected static readonly XmlSerializerNamespaces EmptyNamespace = new XmlSerializerNamespaces(new[] { XmlQualifiedName.Empty });

        protected static readonly XmlWriterSettings OutputXmlSettings = new XmlWriterSettings
        {
            NamespaceHandling = NamespaceHandling.OmitDuplicates,
            Encoding = new UTF8Encoding(false),
            Indent = true
        };

        protected T Cached { get; set; }

        private string _filePath;
        public virtual string FilePath
        {
            get => _filePath ?? Path.Combine(RootPath, DefaultConfigFile);
            set => _filePath = value;
        }

        public abstract T Load(T toDeserialize);

        public abstract void Save(T toSerialize);

        protected static T FailedLoadReturnValue()
        {
            return (T)Convert.ChangeType(null, typeof(T));
        }

        protected static XDocument GetConfigurationDoc(string file)
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
            catch
            {
                // this is fine - we'll just return an empty XDocument.
            }

            output = new XDocument();
            output.Add(new XElement(RootElement));
            return output;
        }

        protected static XElement GetNodeByName(XContainer doc, string name)
        {
            return doc.Descendants().FirstOrDefault(e => e.Name.LocalName.Equals(name));
        }

        protected void EnsurePathExists()
        {
            var folder = Path.GetDirectoryName(FilePath);
            if (folder != null && !Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
        }

        protected T CachedOrNotFound()
        {
            return !File.Exists(FilePath) ? FailedLoadReturnValue() : Cached;
        }
    }
}
