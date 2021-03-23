using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Castle.Core;
using Castle.MicroKernel;
using Castle.MicroKernel.ModelBuilder;
using Rubberduck.UI;
using Rubberduck.UI.Command;

namespace Rubberduck.Root
{
    // Loosely based on https://github.com/castleproject/Windsor/blob/36fbebd9a471f88b43044f39704dc5f19e669e6f/src/Castle.Windsor/MicroKernel/ModelBuilder/Inspectors/PropertiesDependenciesModelInspector.cs
    internal class RubberduckViewModelPropertiesInspector : IContributeComponentModelConstruction
    {
        public void ProcessModel(IKernel kernel, ComponentModel model)
        {
            var targetType = model.Implementation;

            // we only inject properties on ViewModels
            if (!(targetType.Name.EndsWith("ViewModel") && targetType.IsBasedOn(typeof(ViewModelBase))))
            {
                return;
            }

            var properties = GetProperties(model, targetType);
           
            foreach (var property in properties)
            {
                model.AddProperty(BuildDependency(property));
            }
        }

        private PropertySet BuildDependency(PropertyInfo property)
        {
            var dependency = new PropertyDependencyModel(property, isOptional: false);
            return new PropertySet(property, dependency);
        }

        private IEnumerable<PropertyInfo> GetProperties(ComponentModel model, Type targetType)
        {
            const BindingFlags bindingFlags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly;
            return targetType.GetProperties(bindingFlags).ToList()
                .Where(property => property.CanWrite
                                   && property.GetSetMethod() != null
                                   && property.PropertyType.IsBasedOn(typeof(CommandBase))
                                   && !property.PropertyType.IsAbstract);
        }
    }
}
