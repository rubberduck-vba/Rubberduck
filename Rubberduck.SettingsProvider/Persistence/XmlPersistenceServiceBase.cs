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
    internal abstract class XmlPersistenceServiceBase<T> : IFilePersistenceService<T> where T : class, IEquatable<T>, new()
    {
        private const string DefaultConfigFile = "rubberduck.config";

        protected readonly string RootPath;
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

        protected XmlPersistenceServiceBase(IPersistencePathProvider pathProvider)
        {
            RootPath = pathProvider.DataRootPath;
        }

        private string _filePath;
        public virtual string FilePath
        {
            get => _filePath ?? Path.Combine(RootPath, DefaultConfigFile);
            set => _filePath = value;
        }

        public abstract T Load(T toDeserialize, string nonDefaultFilePath = null);

        public abstract void Save(T toSerialize, string nonDefaultFilePath = null);

        protected static T FailedLoadReturnValue()
        {
            return null;
        }

        protected static XDocument GetConfigurationDoc(string file)
        {
            XDocument output;
            if (File.Exists(file))
            {
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
            }

            output = new XDocument();
            output.Add(new XElement(RootElement));
            return output;
        }

        protected static XElement GetNodeByName(XContainer doc, string name)
        {
            return doc.Descendants().FirstOrDefault(e => e.Name.LocalName.Equals(name));
        }

        protected void EnsurePathExists(string nonDefaultFilePath = null)
        {
            var filePath = string.IsNullOrWhiteSpace(nonDefaultFilePath) ? FilePath : nonDefaultFilePath;
            var folder = Path.GetDirectoryName(filePath);
            if (folder != null && !Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
        }
    }
}
