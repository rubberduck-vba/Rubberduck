using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Templates
{
    public interface ITemplateFileHandlerProvider
    {
        ITemplateFileHandler CreateTemplateFileHandler(string templateName);
        IEnumerable<string> GetTemplateNames();
    }

    public class TemplateFileHandlerProvider : ITemplateFileHandlerProvider
    {
        private readonly string _rootPath;

        public TemplateFileHandlerProvider(IPersistencePathProvider pathProvider)
        {
            _rootPath = pathProvider.DataFolderPath("Templates");
        }

        public ITemplateFileHandler CreateTemplateFileHandler(string templateName)
        {
            if (!Directory.Exists(_rootPath))
            {
                Directory.CreateDirectory(_rootPath);
            }

            var fullPath = Path.Combine(_rootPath, templateName);
            if (!Directory.Exists(Path.GetDirectoryName(fullPath)))
            {
                throw new InvalidOperationException("Cannot provide a path for where the parent directory do not exist");
            }
            return  new TemplateFileHandler(fullPath);
        }

        public IEnumerable<string> GetTemplateNames()
        {
            var info = new DirectoryInfo(_rootPath);
            return info.GetFiles().Select(file => file.Name).ToList();
        }
    }

    public interface ITemplateFileHandler
    {
        bool Exists { get; }
        string Read();
        void Write(string content);
    }
    
    public class TemplateFileHandler : ITemplateFileHandler
    {
        private readonly string _fullPath;

        public TemplateFileHandler(string fullPath)
        {
            _fullPath = fullPath + (fullPath.EndsWith(".rdt") ? string.Empty : ".rdt");
        }

        public bool Exists => File.Exists(_fullPath);

        public string Read()
        {
            return Exists ? File.ReadAllText(_fullPath) : null;
        }

        public void Write(string content)
        {
            File.WriteAllText(_fullPath, content);
        }
    }
}