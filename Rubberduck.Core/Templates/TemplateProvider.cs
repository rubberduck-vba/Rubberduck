using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;

namespace Rubberduck.Templates
{
    public interface ITemplateProvider
    {
        ITemplate Load(string templateName);
        ObservableCollection<Template> GetTemplates();
    }

    public class TemplateProvider : ITemplateProvider
    {
        private readonly ITemplateFileHandlerProvider _provider;

        public TemplateProvider(ITemplateFileHandlerProvider provider)
        {
            _provider = provider;
        }

        public ITemplate Load(string templateName)
        {
            var handler = _provider.CreateTemplateFileHandler(templateName);
            return new Template(templateName, handler);
        }

        private Lazy<ObservableCollection<Template>> LazyList => new Lazy<ObservableCollection<Template>>(() =>
        {
            var list = new ObservableCollection<Template>();
            foreach (var templateName in _provider.GetTemplateNames())
            {
                var handler = _provider.CreateTemplateFileHandler(templateName);
                list.Add(new Template(templateName, handler));
            }

            var set = Resources.Templates.ResourceManager.GetResourceSet(CultureInfo.CurrentUICulture, true, true);
            foreach (DictionaryEntry entry in set)
            {
                const string NameSuffix = "_Name";
                var key = (string)entry.Key;
                if (key.EndsWith(NameSuffix))
                {
                    var templateName = key.Substring(0, key.Length - NameSuffix.Length);
                    var handler = _provider.CreateTemplateFileHandler(templateName);
                    if (!list.Any(t => t.Name == templateName))
                    {
                        list.Add(new Template(templateName, handler));
                    }
                }
            }

            return list;
        });

        public ObservableCollection<Template> GetTemplates() => LazyList.Value;
    }
}
