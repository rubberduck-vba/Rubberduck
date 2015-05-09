using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml.Serialization;

namespace Rubberduck.Config
{
    public abstract class XmlConfigurationServiceBase<T> : IConfigurationService<T>
    {
        /// <summary>
        /// Defines the root path where all Rubberduck Configuration files are stored.
        /// </summary>
        protected readonly string rootPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck");

        /// <summary>
        /// Implementation of this should combine the rootPath with the xml file name of the object to be serialized.
        /// </summary>
        protected abstract string ConfigFile { get; }

        /// <summary>
        /// Serializes the configuration object to an XML file.
        /// </summary>
        /// <param name="toSerialize">The Configuration Object to be serialized and saved.</param>
        public void SaveConfiguration(T toSerialize)
        {
            var folder = Path.GetDirectoryName(ConfigFile);
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }

            var serializer = new XmlSerializer(typeof(T));
            using (var writer = new StreamWriter(ConfigFile))
            {
                serializer.Serialize(writer, toSerialize);
            }
        }

        /// <summary>
        /// Reads and de-serializes an xml conofiguration file.
        /// </summary>
        /// <returns>Configuration object of type <typeparamref name="T"/></returns>
        public virtual T LoadConfiguration()
        {
            try
            {
                using (var reader = new StreamReader(ConfigFile))
                {
                    var deserializer = new XmlSerializer(typeof(T));
                    return (T)deserializer.Deserialize(reader);
                }
            }
            catch (IOException ex)
            {
                return HandleIOException(ex);
            }
            catch (InvalidOperationException ex)
            {
                return HandleInvalidOperationException(ex);
            }

        }

        /// <summary>
        /// Defines the action, if any, to be taken if an IOException occurs when trying to load a configuration.
        /// </summary>
        protected abstract T HandleIOException(IOException ex);

        /// <summary>
        /// Defines the action, if any, to be taken if an InvalidOperationException occurs when trying to load a configuration.
        /// </summary>
        /// <param name="ex"></param>
        protected abstract T HandleInvalidOperationException(InvalidOperationException ex);

    }
}
