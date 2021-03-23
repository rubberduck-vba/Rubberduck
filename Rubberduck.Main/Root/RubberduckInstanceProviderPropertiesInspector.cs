using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Castle.Core;
using Castle.MicroKernel;
using Castle.MicroKernel.ModelBuilder;

namespace Rubberduck.Root
{
    internal class RubberduckInstanceProviderPropertiesInspector : IContributeComponentModelConstruction
    {
        public void ProcessModel(IKernel kernel, ComponentModel model)
        {
            var targetType = model.Implementation;

            if (!(targetType == typeof(InstanceProvider)))
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
            const BindingFlags bindingFlags = BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly;
            return targetType.GetProperties(bindingFlags).ToList()
                .Where(property => property.CanWrite
                                   && property.GetSetMethod(true) != null
                                   && !property.PropertyType.IsAbstract
                                   && property.Name.EndsWith("Instance"));
        }
    }
}
