using System;
using System.IO;
using System.Runtime.Serialization;
using System.Xml;

namespace Rubberduck.SettingsProvider
{
    internal class XmlContractPersistenceService<T> : XmlPersistenceServiceBase<T> where T : class, IEquatable<T>, new()
    {
        private const string DefaultConfigFile = "rubberduck.references";

        // ReSharper disable once StaticMemberInGenericType
        private static readonly DataContractSerializerSettings SerializerSettings = new DataContractSerializerSettings
        {
            RootNamespace = XmlDictionaryString.Empty            
        };

        public XmlContractPersistenceService(IPersistencePathProvider pathProvider) : base(pathProvider) { }

        private string _filePath;
        public override string FilePath
        {
            get => _filePath ?? Path.Combine(RootPath, DefaultConfigFile);
            set => _filePath = value;
        }

        public override T Load(T toDeserialize, string nonDefaultFilePath = null)
        {
            var filePath = string.IsNullOrWhiteSpace(nonDefaultFilePath) ? FilePath : nonDefaultFilePath;
            if (!File.Exists(filePath))
            {
                return Cached;
            }
            
            try
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                using (var reader = XmlReader.Create(stream))
                {
                    var serializer = new DataContractSerializer(typeof(T));
                    return (T)serializer.ReadObject(reader);
                }
            }
            catch(Exception ex)
            {
                return FailedLoadReturnValue();
            }
        }

        public override void Save(T toSerialize, string nonDefaultFilePath = null)
        {
            var filePath = string.IsNullOrWhiteSpace(nonDefaultFilePath) ? FilePath : nonDefaultFilePath;
            EnsurePathExists(filePath);

            using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            using (var writer = XmlWriter.Create(stream, OutputXmlSettings))
            {
                var serializer = new DataContractSerializer(typeof(T), SerializerSettings);
                serializer.WriteObject(writer, toSerialize);

                Cached = toSerialize;
            }
        }
    }
}
