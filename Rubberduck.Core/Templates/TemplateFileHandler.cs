using System;
using System.Collections.Generic;
using System.IO.Abstractions;
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
        private readonly Lazy<string> _rootPath;
        private readonly IFileSystem _filesystem;

        public TemplateFileHandlerProvider(
            IPersistencePathProvider pathProvider,
            IFileSystem fileSystem)
        {
            _rootPath = new Lazy<string>(() => pathProvider.DataFolderPath("Templates"));
            _filesystem = fileSystem;
        }

        public ITemplateFileHandler CreateTemplateFileHandler(string templateName)
        {
            EnsureRootPathExists();
            var fullPath = _filesystem.Path.Combine(_rootPath.Value, templateName);
            if (!_filesystem.Directory.Exists(_filesystem.Path.GetDirectoryName(fullPath)))
            {
                throw new InvalidOperationException("Cannot provide a path for where the parent directory do not exist");
            }
            return  new TemplateFileHandler(fullPath, _filesystem);
        }

        public IEnumerable<string> GetTemplateNames()
        {
            EnsureRootPathExists();
            var info = _filesystem.DirectoryInfo.FromDirectoryName(_rootPath.Value);
            return info.GetFiles().Select(file => file.Name).ToList();
        }

        private void EnsureRootPathExists()
        {
            if (!_filesystem.Directory.Exists(_rootPath.Value))
            {
                _filesystem.Directory.CreateDirectory(_rootPath.Value);
            }
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
        private readonly IFileSystem _filesystem;

        public TemplateFileHandler(
            string fullPath,
            IFileSystem fileSystem)
        {
            _fullPath = fullPath + (fullPath.EndsWith(Template.TemplateExtension) ? string.Empty : Template.TemplateExtension);
            _filesystem = fileSystem;
        }

        public bool Exists => _filesystem.File.Exists(_fullPath);

        public string Read()
        {
            return Exists ? _filesystem.File.ReadAllText(_fullPath) : null;
        }

        public void Write(string content)
        {
            _filesystem.File.WriteAllText(_fullPath, content);
        }
    }
}