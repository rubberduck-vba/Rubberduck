using System;
using System.Linq;
using System.Runtime.InteropServices;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Application
{
    [ComVisible(false)]
    public abstract class HostApplicationBase<TApplication> : IHostApplication
        where TApplication : class
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        protected readonly TApplication Application;
        protected HostApplicationBase(string applicationName)
        {
            ApplicationName = applicationName;

            try
            {
                Application = (TApplication)Marshal.GetActiveObject($"{applicationName}.Application");
            }
            catch (COMException)
            {
                Application = null; // unit tests don't need it anyway.
            }
        }

        protected HostApplicationBase(IVBE vbe, string applicationName)
        {
            ApplicationName = applicationName;

            try
            {
                var appProperty = ApplicationPropertyFromDocumentModule(vbe);
                if (appProperty != null)
                {
                    Application = (TApplication)appProperty.Object;
                }
                else
                {
                    Application = (TApplication)Marshal.GetActiveObject($"{applicationName}.Application");
                }
                    
            }
            catch (COMException)
            {
                Application = null; // unit tests don't need it anyway.
            }
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

        ~HostApplicationBase()
        {
			Dispose(false);
        }

        public string ApplicationName { get; }

        public abstract void Run(dynamic declaration);

        public virtual object Run(string name, params object[] args)
        {
            return null;
        }

        private int? _rcwReferenceCount;
        public void Release(bool final = false)
        {
            if (HasBeenReleased)
            {
                _logger.Warn($"Tried to release an application object type {this.GetType()} that had already been released.");
                return;
            }
            if (Application == null)
            {
                _rcwReferenceCount = 0;
                _logger.Warn($"Tried to release an application object that was null.");
                return;
            }

            if (!Marshal.IsComObject(Application))
            {
                _rcwReferenceCount = 0;
                _logger.Warn($"Tried to release an application objects of type {this.GetType()} that is not a COM object.");
                return;
            }

            try
            {
                if (final)
                {
                    _rcwReferenceCount = Marshal.FinalReleaseComObject(Application);
                    if (HasBeenReleased)
                    {
                        _logger.Trace($"Final released application object of type {this.GetType()}.");
                    }
                    else
                    {
                        _logger.Warn($"Final released application object of type {this.GetType()} did not release the object: remaining reference count is {_rcwReferenceCount}.");
                    }
                }
                else
                {
                    _rcwReferenceCount = Marshal.ReleaseComObject(Application);
                    if (_rcwReferenceCount >= 0)
                    {
                        _logger.Trace($"Released application object of type {this.GetType()} with remaining reference count {_rcwReferenceCount}.");
                    }
                    else
                    {
                        _logger.Warn($"Released application object of type {this.GetType()} whose underlying RCW has already been released from outside the SafeComWrapper.");
                    }
                }
            }
            catch (COMException exception)
            {
                var logMessage = $"Failed to release application object of type {this.GetType()}.";
                if (_rcwReferenceCount.HasValue)
                {
                    logMessage = logMessage + $"The previous reference count has been {_rcwReferenceCount}.";
                }
                else
                {
                    logMessage = logMessage + "There has not yet been an attempt to release the application object.";
                }

                _logger.Warn(exception, logMessage);
            }
        }

        public bool HasBeenReleased => _rcwReferenceCount <= 0;

        public void Dispose()
        {
            Dispose(true);
			GC.SuppressFinalize(this);
        }

		private bool _disposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }
			
			// clean up managed resources
			if (Application != null && !HasBeenReleased)
            {
                Release();
            }
		
            if (disposing) 
			{ 
				// we don't have any managed resources to clean up right now.
			}

			_disposed = true;
        }
    }
}
