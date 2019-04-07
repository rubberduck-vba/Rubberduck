using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Rubberduck.Templates
{
    public interface ITemplateProvider
    {
        ITemplate Load(string templateName);
        IEnumerable<Template> GetTemplates();
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

        private Lazy<IEnumerable<Template>> LazyList => new Lazy<IEnumerable<Template>>(() =>
        {
            var list = new List<Template>();
            var set = Rubberduck.Resources.Templates.ResourceManager.GetResourceSet(CultureInfo.CurrentUICulture, true, true);

            foreach (DictionaryEntry entry in set)
            {
                var key = (string)entry.Key;
                if (key.EndsWith("_Name"))
                {
                    var templateName = key.Substring(0, key.Length - "_Name".Length);
                    var handler = _provider.CreateTemplateFileHandler(templateName);
                    list.Add(new Template(templateName, handler));
                }
            }

            foreach (var templateName in _provider.GetTemplateNames())
            {
                if (list.Any(e => e.Name == templateName))
                {
                    continue;
                }

                var handler = _provider.CreateTemplateFileHandler(templateName);
                list.Add(new Template(templateName, handler));
            }

            return list;
        });

        public IEnumerable<Template> GetTemplates() => LazyList.Value;
    }
}
