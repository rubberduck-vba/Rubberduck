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
    internal abstract class XmlPersistenceServiceBase<T> : IPersistenceService<T> where T : class, IEquatable<T>, new()
    {
        protected readonly string RootPath;
        protected const string RootElement = "Configuration";

        protected static readonly XmlSerializerNamespaces EmptyNamespace = new XmlSerializerNamespaces(new[] { XmlQualifiedName.Empty });
        protected static readonly UTF8Encoding OutputEncoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);        
        protected static readonly XmlWriterSettings OutputXmlSettings = new XmlWriterSettings
        {
            NamespaceHandling = NamespaceHandling.OmitDuplicates,
            Encoding = OutputEncoding,
            Indent = true
        };

        protected XmlPersistenceServiceBase(IPersistencePathProvider pathProvider)
        {
            RootPath = pathProvider.DataRootPath;
        }
        
        protected abstract string FilePath { get; }

        public T Load(string path = default)
        {
            return Read(string.IsNullOrEmpty(path) ? FilePath : path);
        }

        public void Save(T toSerialize, string path = default)
        {
            var targetPath = string.IsNullOrEmpty(path) ? FilePath : path;
            EnsureDirectoryExists(targetPath);
            Write(toSerialize, targetPath);
        }

        protected abstract T Read(string path);
        protected abstract void Write(T toSerialize, string path);

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

        protected void EnsureDirectoryExists(string filePath)
        {
            var folder = Path.GetDirectoryName(filePath);
            if (folder != null && !Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
        }
    }
}
