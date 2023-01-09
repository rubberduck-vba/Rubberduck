/**
Shamelessly stolen from https://github.com/dotnet/Nerdbank.GitVersioning/blob/3e4e1f8249ba70fd576b524ce12398ee398884fc/src/Nerdbank.GitVersioning.Tasks/ContextAwareTask.cs
*/

namespace Rubberduck.Deployment.Build
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Reflection;
#if NETCOREAPP
    using System.Runtime.Loader;
#endif
    using Microsoft.Build.Framework;
    using Microsoft.Build.Utilities;
#if NETCOREAPP
    using Nerdbank.GitVersioning;
#endif

    public abstract class ContextAwareTask : AppDomainIsolatedTask
    {
        protected virtual string ManagedDllDirectory => Path.GetDirectoryName(new Uri(this.GetType().GetTypeInfo().Assembly.CodeBase).LocalPath);

        protected virtual string UnmanagedDllDirectory => null;

        public override bool Execute()
        {
#if NETCOREAPP
            string taskAssemblyPath = new Uri(this.GetType().GetTypeInfo().Assembly.CodeBase).LocalPath;

            Assembly inContextAssembly = GitLoaderContext.Instance.LoadFromAssemblyPath(taskAssemblyPath);
            Type innerTaskType = inContextAssembly.GetType(this.GetType().FullName);
            object innerTask = Activator.CreateInstance(innerTaskType);

            var outerProperties = this.GetType().GetRuntimeProperties().ToDictionary(i => i.Name);
            var innerProperties = innerTaskType.GetRuntimeProperties().ToDictionary(i => i.Name);
            var propertiesDiscovery = from outerProperty in outerProperties.Values
                                      where outerProperty.SetMethod != null && outerProperty.GetMethod != null
                                      let innerProperty = innerProperties[outerProperty.Name]
                                      select new { outerProperty, innerProperty };
            var propertiesMap = propertiesDiscovery.ToArray();
            var outputPropertiesMap = propertiesMap.Where(pair => pair.outerProperty.GetCustomAttribute<OutputAttribute>() != null).ToArray();

            foreach (var propertyPair in propertiesMap)
            {
                object outerPropertyValue = propertyPair.outerProperty.GetValue(this);
                propertyPair.innerProperty.SetValue(innerTask, outerPropertyValue);
            }

            var executeInnerMethod = innerTaskType.GetMethod(nameof(ExecuteInner), BindingFlags.Instance | BindingFlags.NonPublic);
            bool result = (bool)executeInnerMethod.Invoke(innerTask, new object[0]);

            foreach (var propertyPair in outputPropertiesMap)
            {
                propertyPair.outerProperty.SetValue(this, propertyPair.innerProperty.GetValue(innerTask));
            }

            return result;
#else
            // On .NET Framework (on Windows), we find native binaries by adding them to our PATH.
            if (this.UnmanagedDllDirectory != null)
            {
                string pathEnvVar = Environment.GetEnvironmentVariable("PATH");
                string[] searchPaths = pathEnvVar.Split(Path.PathSeparator);
                if (!searchPaths.Contains(this.UnmanagedDllDirectory, StringComparer.OrdinalIgnoreCase))
                {
                    pathEnvVar += Path.PathSeparator + this.UnmanagedDllDirectory;
                    Environment.SetEnvironmentVariable("PATH", pathEnvVar);
                }
            }

            return this.ExecuteInner();
#endif
        }

        protected abstract bool ExecuteInner();
    }
}