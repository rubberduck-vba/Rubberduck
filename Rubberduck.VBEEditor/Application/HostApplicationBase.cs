using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Application
{
    [ComVisible(false)]
    public abstract class HostApplicationBase<TApplication> : SafeComWrapper<TApplication>, IHostApplication
        where TApplication : class
    {
        protected HostApplicationBase(string applicationName)
        :base(ApplicationFomComReflection(applicationName))
        {
            ApplicationName = applicationName;
        }

        protected HostApplicationBase(IVBE vbe, string applicationName)
        :base(ApplicationFomVbe(vbe, applicationName))
        {
            ApplicationName = applicationName;
        }

        private static TApplication ApplicationFomComReflection(string applicationName)
        {
            TApplication application;
            try
            {
                application = (TApplication)Marshal.GetActiveObject($"{applicationName}.Application");
            }
            catch (COMException)
            {
                application = null; // unit tests don't need it anyway.
            }
            return application;
        }

        private static TApplication ApplicationFomVbe(IVBE vbe, string applicationName)
        {
            TApplication application;
            try
            {
                var appProperty = ApplicationPropertyFromDocumentModule(vbe);
                if (appProperty != null)
                {
                    application = (TApplication)appProperty.Object;
                }
                else
                {
                    application = (TApplication)Marshal.GetActiveObject($"{applicationName}.Application");
                }

            }
            catch (COMException)
            {
                application = null; // unit tests don't need it anyway.
            }
            return application;
        }

        private static IProperty ApplicationPropertyFromDocumentModule(IVBE vbe)
        {
            using (var projects = vbe.VBProjects)
            {
                foreach (var project in projects)
                {
                    try
                    {
                        if (project.Protection == ProjectProtection.Locked)
                        {
                            continue;
                        }
                        using (var components = project.VBComponents)
                        {
                            foreach (var component in components)
                            {
                                try
                                {
                                    if (component.Type != ComponentType.Document)
                                    {
                                        continue;
                                    }
                                    using (var properties = component.Properties)
                                    {
                                        if (properties.Count <= 1)
                                        {
                                            continue;
                                        }
                                        foreach (var property in properties)
                                        {
                                            if (property.Name == "Application")
                                            {
                                                return property;
                                            }
                                            property.Dispose();
                                        }
                                    }
                                }
                                finally
                                {
                                    component.Dispose();
                                }
                            }
                        }
                    }
                    finally
                    {
                        project?.Dispose();
                    }
                }
                return null;
            }
        }

        protected TApplication Application => Target;

        public string ApplicationName { get; }

        public abstract void Run(dynamic declaration);

        public virtual object Run(string name, params object[] args)
        {
            return null;
        }

        public override bool Equals(ISafeComWrapper<TApplication> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : HashCode.Compute(Target);
        }

        ~HostApplicationBase()
        {
            Dispose(false);
        }
    }
}
