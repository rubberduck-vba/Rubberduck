using System;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using System.Linq;

namespace Rubberduck.VBEditor.VBEHost
{
    [ComVisible(false)]
    public abstract class HostApplicationBase<TApplication> : IHostApplication
        where TApplication : class
    {
        private readonly string _applicationName;
        protected readonly TApplication Application;
        protected HostApplicationBase(string applicationName)
        {
            _applicationName = applicationName;

            try
            {
                Application = (TApplication)Marshal.GetActiveObject(applicationName + ".Application");
            }
            catch (COMException)
            {
                Application = null; // unit tests don't need it anyway.
            }
        }

        protected HostApplicationBase(VBE vbe, string applicationName)
        {
            _applicationName = applicationName;

            try
            {
                var appProperty = vbe.VBProjects
                    .Cast<VBProject>()
                    .Where(project => project.Protection == vbext_ProjectProtection.vbext_pp_none)
                    .SelectMany(project => project.VBComponents.Cast<VBComponent>())
                    .Where(component => component.Type == vbext_ComponentType.vbext_ct_Document
                    && component.Properties.Count > 1)
                    .SelectMany(component => component.Properties.OfType<Property>())
                    .FirstOrDefault(property => property.Name == "Application");
                if (appProperty != null)
                {
                    Application = (TApplication)appProperty.Object;
                }
                else
                {
                    Application = (TApplication)Marshal.GetActiveObject(applicationName + ".Application");
                }
                    
            }
            catch (COMException)
            {
                Application = null; // unit tests don't need it anyway.
            }
        }

        ~HostApplicationBase()
        {
			Dispose(false);
        }

        public string ApplicationName
        {
            get { return _applicationName; }
        }

        public abstract void Run(QualifiedMemberName qualifiedMemberName);

        public void Dispose()
        {
            Dispose(true);
			GC.SuppressFinalize(this);
        }

		private bool _disposed;
        protected virtual void Dispose(bool disposing)
        {
			if (_disposed) { return; }
			
			// clean up managed resources
			if (Application != null)
            {
                Marshal.ReleaseComObject(Application);
            }
		
            if (disposing) 
			{ 
				// we don't have any managed resources to clean up right now.
			}

			_disposed = true;
        }
    }
}
