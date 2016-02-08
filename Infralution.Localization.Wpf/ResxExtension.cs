//
//      FILE:   ResxExtension.cs.
//
// COPYRIGHT:   Copyright 2008 
//              Infralution
//
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Resources;
using System.Reflection;
using System.Windows;
using System.ComponentModel;
using System.Globalization;
using System.Threading;
using System.Diagnostics;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Drawing;
using System.Windows.Interop;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Data;
using System.Windows.Threading;
using System.Collections.ObjectModel;
using System.Management;
using Microsoft.Win32;

[assembly: XmlnsDefinition("http://schemas.microsoft.com/winfx/2006/xaml/presentation", "Infralution.Localization.Wpf")]
[assembly: XmlnsDefinition("http://schemas.microsoft.com/winfx/2007/xaml/presentation", "Infralution.Localization.Wpf")]
[assembly: XmlnsDefinition("http://schemas.microsoft.com/winfx/2008/xaml/presentation", "Infralution.Localization.Wpf")]

namespace Infralution.Localization.Wpf
{
    /// <summary>
    /// Defines the handling method for the <see cref="ResxExtension.GetResource"/> event
    /// </summary>
    /// <param name="resxName">The name of the resx file</param>
    /// <param name="key">The resource key within the file</param>
    /// <param name="culture">The culture to get the resource for</param>
    /// <returns>The resource</returns>
    public delegate object GetResourceHandler(string resxName, string key, CultureInfo culture);

    /// <summary>
    /// A markup extension to allow resources for WPF Windows and controls to be retrieved
    /// from an embedded resource (resx) file associated with the window or control
    /// </summary>
    /// <remarks>
    /// Supports design time switching of the Culture via a Tray Notification icon. Loading
    /// of culture specific satellite assemblies at design time (within the XDesProc designer
    /// process) is done by probing the sub-directories associated with the running Visual Studio 
    /// hosting process (*.vshost) for the latest matching assembly.  If you have disabled the
    /// hosting process or are using Expression Blend then you can set the sub-directories to
    /// search at design time by creating a string Value in the registry:
    /// 
    /// HKEY_CURRENT_USER\Software\ResxExtension\AssemblyPath
    /// 
    /// and set the value to a semi-colon delimited list of directories to search  
    /// </remarks>
    [MarkupExtensionReturnType(typeof(object))]
    [ContentProperty("Children")]
    public class ResxExtension : ManagedMarkupExtension
    {
       
        #region Member Variables

        /// <summary>
        /// The explicitly set embedded Resx Name (if any)
        /// </summary>
        private string _resxName;

        /// <summary>
        /// The default resx name (based on the attached property)
        /// </summary>
        private string _defaultResxName;

        /// <summary>
        /// The key used to retrieve the resource
        /// </summary>
        private string _key;

        /// <summary>
        /// The default value for the property
        /// </summary>
        private string _defaultValue;

        /// <summary>
        /// The key used to retrieve the BindingTargetNullValue
        /// </summary>
        private string _bindingTargetNullKey;

        /// <summary>
        /// The resource manager to use for this extension.  Holding a strong reference to the
        /// Resource Manager keeps it in the cache while ever there are ResxExtensions that
        /// are using it.
        /// </summary>
        private ResourceManager _resourceManager;

        /// <summary>
        /// The binding (if any) used to store the binding properties for the extension  
        /// </summary>
        private Binding _binding;

        /// <summary>
        /// The child ResxExtensions (if any) when using MultiBinding expressions
        /// </summary>
        private Collection<ResxExtension> _children = new Collection<ResxExtension>();

        /// <summary>
        /// Cached resource managers
        /// </summary>
        private static Dictionary<string, WeakReference> _resourceManagers = new Dictionary<string, WeakReference>();

        /// <summary>
        /// The manager for resx extensions
        /// </summary>
        private static MarkupExtensionManager _markupManager = new MarkupExtensionManager(40);

        /// <summary>
        /// The directories to probe for satellite assemblies when running inside the Visual Studio
        /// designer process (XDesProc)
        /// </summary>
        private static List<string> _assemblyProbingPaths;

        #endregion

        #region Public Interface


        /// <summary>
        /// This event allows a designer or preview application (such as Globalizer.NET) to
        /// intercept calls to get resources and provide the values instead dynamically
        /// </summary>
        public static event GetResourceHandler GetResource;

        /// <summary>
        /// Create a new instance of the markup extension
        /// </summary>
        public ResxExtension()
            : base(_markupManager)
        {
        }

        /// <summary>
        /// Create a new instance of the markup extension
        /// </summary>
        /// <param name="key">The key used to get the value from the resources</param>
        public ResxExtension(string key)
            : base(_markupManager)
        {
            this._key = key;
        }

        /// <summary>
        /// The fully qualified name of the embedded resx (without .resources) to get
        /// the resource from
        /// </summary>
        public string ResxName
        {
            get 
            {
                // if the ResxName property is not set explicitly then check the attached property
                //
                string result = _resxName;
                if (string.IsNullOrEmpty(result))
                {
                    if (_defaultResxName == null)
                    {
                        WeakReference targetRef = TargetObjects.Find(target => target.IsAlive);
                        if (targetRef != null)
                        {
                            if (targetRef.Target is DependencyObject)
                            {
                               _defaultResxName = (targetRef.Target as DependencyObject).GetValue(DefaultResxNameProperty) as string;
                            }
                        }
                    }
                    result = _defaultResxName;
                }
                return result; 
            }
            set 
            { 
                _resxName = value; 
            }
        }

        /// <summary>
        /// The name of the resource key
        /// </summary>
        public string Key
        {
            get { return _key; }
            set { _key = value; }
        }

        /// <summary>
        /// The default value to use if the resource can't be found
        /// </summary>
        /// <remarks>
        /// This particularly useful for properties which require non-null
        /// values because it allows the page to be displayed even if
        /// the resource can't be loaded
        /// </remarks>
        public string DefaultValue
        {
            get { return _defaultValue; }
            set { _defaultValue = value; }
        }

        /// <summary>
        /// The child Resx elements (if any) 
        /// </summary>
        /// <remarks>
        /// You can nest Resx elements in this case the parent Resx element
        /// value is used as a format string to format the values from child Resx
        /// elements similar to a <see cref="MultiBinding"/> eg If a Resx has two 
        /// child elements then you 
        /// </remarks>
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public Collection<ResxExtension> Children
        {
            get { return _children; }
        }

        #region Delegated Binding properties

        /// <summary>
        /// Return the associated binding for the extension
        /// </summary>
        public Binding Binding
        {
            get
            {
                if (_binding == null)
                {
                    _binding = new Binding();
                }
                return _binding;
            }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.ElementName"/>.
        /// </summary>
        [DefaultValue(null)]
        public string BindingElementName
        {
            get { return Binding.ElementName; }
            set { Binding.ElementName = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.Path"/>.
        /// </summary>
        [DefaultValue(null)]
        public PropertyPath BindingPath
        {
            get { return Binding.Path; }
            set { Binding.Path = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.RelativeSource"/>.
        /// </summary>
        [DefaultValue(null)]
        public RelativeSource BindingRelativeSource
        {
            get { return Binding.RelativeSource; }
            set { Binding.RelativeSource = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.Source"/>.
        /// </summary>
        [DefaultValue(null)]
        public object BindingSource
        {
            get { return Binding.Source; }
            set { Binding.Source = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.XPath"/>.
        /// </summary>
        [DefaultValue(null)]
        public string BindingXPath
        {
            get { return Binding.XPath; }
            set { Binding.XPath = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.Converter"/>.
        /// </summary>
        [DefaultValue(null)]
        public IValueConverter BindingConverter
        {
            get { return Binding.Converter; }
            set { Binding.Converter = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.BindingBase.BindingGroupName"/>.
        /// </summary>
        [DefaultValue(null)]
        public string BindingGroupName
        {
            get { return Binding.BindingGroupName; }
            set { Binding.BindingGroupName = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.ConverterCulture"/>.
        /// </summary>
        [DefaultValue(null)]
        public CultureInfo BindingConverterCulture
        {
            get { return Binding.ConverterCulture; }
            set { Binding.ConverterCulture = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.ConverterParameter"/>.
        /// </summary>
        [DefaultValue(null)]
        public object BindingConverterParameter
        {
            get { return Binding.ConverterParameter; }
            set { Binding.ConverterParameter = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.BindsDirectlyToSource"/>.
        /// </summary>
        [DefaultValue(false)]
        public bool BindsDirectlyToSource
        {
            get { return Binding.BindsDirectlyToSource; }
            set { Binding.BindsDirectlyToSource = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.Mode"/>.
        /// </summary>
        [DefaultValue(BindingMode.Default)]
        public BindingMode BindingMode
        {
            get { return Binding.Mode; }
            set { Binding.Mode = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.NotifyOnSourceUpdated"/>.
        /// </summary>
        [DefaultValue(false)]
        public bool BindingNotifyOnSourceUpdated
        {
            get { return Binding.NotifyOnSourceUpdated; }
            set { Binding.NotifyOnSourceUpdated = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.NotifyOnTargetUpdated"/>.
        /// </summary>
        [DefaultValue(false)]
        public bool BindingNotifyOnTargetUpdated
        {
            get { return Binding.NotifyOnTargetUpdated; }
            set { Binding.NotifyOnTargetUpdated = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.NotifyOnValidationError"/>.
        /// </summary>
        [DefaultValue(false)]
        public bool BindingNotifyOnValidationError
        {
            get { return Binding.NotifyOnValidationError; }
            set { Binding.NotifyOnValidationError = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.AsyncState"/>.
        /// </summary>
        [DefaultValue(null)]
        public object BindingAsyncState
        {
            get { return Binding.AsyncState; }
            set { Binding.AsyncState = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.IsAsync"/>.
        /// </summary>
        [DefaultValue(false)]
        public bool BindingIsAsync
        {
            get { return Binding.IsAsync; }
            set { Binding.IsAsync = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.BindingBase.FallbackValue"/>.
        /// </summary>
        [DefaultValue(null)]
        public object BindingFallbackValue
        {
            get { return Binding.FallbackValue; }
            set { Binding.FallbackValue = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.BindingBase.TargetNullValue"/>.
        /// </summary>
        [DefaultValue(null)]
        public object BindingTargetNullValue
        {
            get { return Binding.TargetNullValue; }
            set { Binding.TargetNullValue = value; }
        }

        /// <summary>
        /// Supply a Resx key to set the BindingTargetNullValue
        /// </summary>
        [DefaultValue(null)]
        public string BindingTargetNullKey
        {
            get { return _bindingTargetNullKey; }
            set 
            {  _bindingTargetNullKey = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.ValidatesOnDataErrors"/>.
        /// </summary>
        [DefaultValue(false)]
        public bool BindingValidatesOnDataErrors
        {
            get { return Binding.ValidatesOnDataErrors; }
            set { Binding.ValidatesOnDataErrors = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.ValidatesOnExceptions"/>.
        /// </summary>
        [DefaultValue(false)]
        public bool BindingValidatesOnExceptions
        {
            get { return Binding.ValidatesOnExceptions; }
            set { Binding.ValidatesOnExceptions = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.UpdateSourceTrigger"/>.
        /// </summary>
        [DefaultValue(UpdateSourceTrigger.Default)]
        public UpdateSourceTrigger BindingUpdateSourceTrigger
        {
            get { return Binding.UpdateSourceTrigger; }
            set { Binding.UpdateSourceTrigger = value; }
        }

        /// <summary>
        /// Use the Resx value to format bound data.  See <see cref="System.Windows.Data.Binding.ValidationRules"/>.
        /// </summary>
        [DefaultValue(false)]
        public Collection<ValidationRule> BindingValidationRules
        {
            get { return Binding.ValidationRules; }
        }

        #endregion

        /// <summary>
        /// Return the value for this instance of the Markup Extension
        /// </summary>
        /// <param name="serviceProvider">The service provider</param>
        /// <returns>The value of the element</returns>
        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            object result = null;

            // register the target and property so we can update them
            //
            RegisterTarget(serviceProvider);

            // Show the icon in the notification tray to allow changing culture at design time
            //
            if (this.IsInDesignMode)
            {
                CultureManager.ShowCultureNotifyIcon();
            }

            if (string.IsNullOrEmpty(Key) && !IsBindingExpression)
                throw new ArgumentException("You must set the resource Key or Binding properties");

            // if the extension is used in a template or as a child of another
            // resx extension (for multi-binding) then return this 
            //
            if (TargetProperty == null || IsMultiBindingChild)
            {
                result = this;
            }
            else
            {
                // if this extension has child Resx elements then invoke AFTER this method has returned
                // to setup the MultiBinding on the target element.  
                //
                if (IsMultiBindingParent)
                {
                    MultiBinding binding = CreateMultiBinding();
                    result = binding.ProvideValue(serviceProvider);
                }
                else if (IsBindingExpression)
                {
                    // if this is a simple binding then return the binding
                    //
                    Binding binding = CreateBinding();
                    result = binding.ProvideValue(serviceProvider);
                }
                else
                {
                    // otherwise return the value from the resources
                    //
                    result = GetValue();
                }
            }
            return result;
        }

        /// <summary>
        /// Return the MarkupManager for this extension
        /// </summary>
        public static MarkupExtensionManager MarkupManager
        {
            get { return _markupManager; }
        }

        /// <summary>
        /// Use the Markup Manager to update all targets
        /// </summary>
        public static void UpdateAllTargets()
        {
            _markupManager.UpdateAllTargets();
        }

        /// <summary>
        /// Update the ResxExtension target with the given key
        /// </summary>
        public static void UpdateTarget(string key)
        {
            foreach (ResxExtension ext in _markupManager.ActiveExtensions)
            {
                if (ext.Key == key)
                {
                    ext.UpdateTargets();
                }
            }
        }

        /// <summary>
        /// The ResxName attached property
        /// </summary>
        public static readonly DependencyProperty DefaultResxNameProperty =
            DependencyProperty.RegisterAttached(
            "DefaultResxName",
            typeof(string),
            typeof(ResxExtension),
            new FrameworkPropertyMetadata(null,
                FrameworkPropertyMetadataOptions.AffectsRender |
                FrameworkPropertyMetadataOptions.Inherits,
                new PropertyChangedCallback(OnDefaultResxNamePropertyChanged)));

        /// <summary>
        /// Get the DefaultResxName attached property for the given target
        /// </summary>
        /// <param name="target">The Target object</param>
        /// <returns>The name of the Resx</returns>
        [AttachedPropertyBrowsableForChildren(IncludeDescendants = true)]
        public static string GetDefaultResxName(DependencyObject target)
        {
            return (string)target.GetValue(DefaultResxNameProperty);
        }

        /// <summary>
        /// Set the DefaultResxName attached property for the given target
        /// </summary>
        /// <param name="target">The Target object</param>
        /// <param name="value">The name of the Resx</param>
        public static void SetDefaultResxName(DependencyObject target, string value)
        {
            target.SetValue(DefaultResxNameProperty, value);
        }

        #endregion

        #region Local Methods

        /// <summary>
        /// Class constructor
        /// </summary>
        static ResxExtension()
        {
            // The Visual Studio 2012/2013 designer process (XDesProc) shadow copies the 
            // assemblies to a cache location.  Unfortunately it doesn't shadow copy the satellite 
            // assemblies - so we have to resolve these ourselves if we want to have support for
            // design time switching of language
            //
            if (AppDomain.CurrentDomain.FriendlyName == "XDesProc.exe")
            {
                _assemblyProbingPaths = new List<string>();

                // check the registry first for a defined assembly path - use OpenBaseKey to avoid Wow64 redirection
                //
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\ResxExtension", false))
                {
                    if (key != null)
                    {
                        string assemblyPath = key.GetValue("AssemblyPath") as string;
                        if (assemblyPath != null)
                        {
                            string[] pathSplit = assemblyPath.Split(';');
                            foreach (string path in pathSplit)
                            {
                                _assemblyProbingPaths.Add(path.Trim());
                            }
                        }
                    }
                }


                // Look for Visual Studio hosting processes and add the path to the probing path - this
                // means that if the hosting process is enabled you don't need to use a registry entry
                //
                foreach (var process in Process.GetProcesses())
                {
                    try
                    {
                        if (process.ProcessName.Contains(".vshost"))
                        {
                            string path = GetProcessFilepath(process.Id);
                            _assemblyProbingPaths.Add(Path.GetDirectoryName(path));
                        }
                    }
                    catch
                    {
                    }
                }
                AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;
            }
        }

        /// <summary>
        /// Return the file path associated with the main module of the given process ID
        /// </summary>
        /// <param name="processId"></param>
        /// <returns></returns>
        private static string GetProcessFilepath(int processId)
        {
            string wmiQueryString = "SELECT ProcessId, ExecutablePath FROM Win32_Process WHERE ProcessId = " + processId;
            using (var searcher = new ManagementObjectSearcher(wmiQueryString))
            {
                using (ManagementObjectCollection results = searcher.Get())
                {
                    foreach (ManagementObject mo in results)
                    {
                        return (string)mo["ExecutablePath"];
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Resolve satellite assemblies when running inside the Visual Studio designer (XDesProc)
        /// process
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        /// <returns>The assembly if found</returns>
        private static Assembly OnAssemblyResolve(object sender, ResolveEventArgs args)
        {
            Assembly result = null;
            string[] nameSplit = args.Name.Split(',');
            if (nameSplit.Length < 3) return null;
            string name = nameSplit[0];

            // Only resolve satellite resource assemblies
            //
            if (!name.EndsWith(".resources")) return null;

            // ignore calls to resolve our own satellite assemblies
            //
            string thisAssembly = Assembly.GetExecutingAssembly().GetName().Name;
            if (name == thisAssembly + ".resources") return null;

            // check that we haven't already loaded the assembly - for some reason AssemblyResolve
            // is still called sometimes after the assembly has already been loaded.  Most recently
            // loaded assemblies are last on the list
            //
            Assembly[] loadedAssemblies = AppDomain.CurrentDomain.GetAssemblies();
            for (int i = loadedAssemblies.Length - 1; i >= 0; i--)
            {
                Assembly assembly = loadedAssemblies[i];
                if (assembly.FullName == args.Name)
                {
                    return assembly;
                }
            }

            // get the culture of the assembly to load
            //
            string[] cultureSplit = nameSplit[2].Split('=');
            if (cultureSplit.Length < 2) return null;
            string culture = cultureSplit[1];

            string fileName = name + ".dll";

            // look for the latest version of the satellite assembly with the given culture 
            // on the assembly probing paths
            //
            string latestFile = null;
            DateTime latestFileTime = DateTime.MinValue;
            foreach (string path in _assemblyProbingPaths)
            {
                string dir = Path.Combine(path, culture);
                string file = Path.Combine(dir, fileName);
                if (File.Exists(file))
                {
                    DateTime fileTime = File.GetLastWriteTime(file);
                    if (fileTime > latestFileTime)
                    {
                        latestFile = file;
                    }
                }
            }

            if (latestFile != null)
            {
                result = Assembly.Load(System.IO.File.ReadAllBytes(latestFile));
            }
            return result;
        }

        /// <summary>
        /// Convert a culture name to a CultureInfo - without exceptions if the name is bad
        /// </summary>
        /// <param name="name">The name of the culture</param>
        /// <returns>The culture if the name was valid, or else null</returns>
        /// <remarks>The CultureInfo constructor throws an exception</remarks>
        static private CultureInfo GetCulture(string name)
        {
            CultureInfo result = null;
            try
            {
                result = new CultureInfo(name);
            }
            catch
            {
            }
            return result;
        }

        /// <summary>
        /// Return a list of the current design time cultures
        /// </summary>
        /// <returns></returns>
        static internal List<CultureInfo> GetDesignTimeCultures()
        {
            List<CultureInfo> result = new List<CultureInfo>();
            if (_assemblyProbingPaths != null)
            {
                foreach (string path in _assemblyProbingPaths)
                {
                    string[] subDirectories = Directory.GetDirectories(path);
                    CultureInfoConverter converter = new CultureInfoConverter();
                    foreach (string subDirectory in subDirectories)
                    {
                        CultureInfo culture = GetCulture(Path.GetFileName(subDirectory));
                        if (culture != null)
                        {
                            result.Add(culture);
                        }
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Create a binding for this Resx Extension
        /// </summary>
        /// <returns>A binding for this Resx Extension</returns>
        private Binding CreateBinding()
        {
            Binding binding = new Binding();
            if (IsBindingExpression)
            {
                // copy all the properties of the binding to the new binding
                //
                if (_binding.ElementName != null)
                {
                    binding.ElementName = _binding.ElementName;
                }
                if (_binding.RelativeSource != null)
                {
                    binding.RelativeSource = _binding.RelativeSource;
                }
                if (_binding.Source != null)
                {
                    binding.Source = _binding.Source;
                }
                
                binding.AsyncState = _binding.AsyncState;
                binding.BindingGroupName = _binding.BindingGroupName;
                binding.BindsDirectlyToSource = _binding.BindsDirectlyToSource;
                binding.Converter = _binding.Converter;
                binding.ConverterCulture = _binding.ConverterCulture;
                binding.ConverterParameter = _binding.ConverterParameter;
                binding.FallbackValue = _binding.FallbackValue;
                binding.IsAsync = _binding.IsAsync;
                binding.Mode = _binding.Mode;
                binding.NotifyOnSourceUpdated = _binding.NotifyOnSourceUpdated;
                binding.NotifyOnTargetUpdated = _binding.NotifyOnTargetUpdated;
                binding.NotifyOnValidationError = _binding.NotifyOnValidationError;
                binding.Path = _binding.Path;
                if (string.IsNullOrEmpty(_bindingTargetNullKey))
                {
                    binding.TargetNullValue = _binding.TargetNullValue;
                }
                else
                {
                    binding.TargetNullValue = GetLocalizedResource(_bindingTargetNullKey);
                }
                binding.UpdateSourceTrigger = _binding.UpdateSourceTrigger;
                binding.ValidatesOnDataErrors = _binding.ValidatesOnDataErrors;
                binding.ValidatesOnExceptions = _binding.ValidatesOnExceptions;
                foreach (ValidationRule rule in _binding.ValidationRules)
                {
                    binding.ValidationRules.Add(rule);
                }
                binding.XPath = _binding.XPath;
                binding.StringFormat = GetValue() as string;
            }
            else
            {
                binding.Source = GetValue();
            }
            return binding;
        }

        /// <summary>
        /// Create new MultiBinding that binds to the child Resx Extensioins
        /// </summary>
        /// <returns></returns>
        private MultiBinding CreateMultiBinding()
        {
            MultiBinding result = new MultiBinding();
            foreach (ResxExtension child in _children)
            {
                // ensure the child has a resx name
                //
                if (child.ResxName == null)
                {
                    child.ResxName = ResxName;
                }
                result.Bindings.Add(child.CreateBinding());
            }
            result.StringFormat = GetValue() as string;
            return result;
        }

        /// <summary>
        /// Have any of the binding properties been set
        /// </summary>
        private bool IsBindingExpression
        {
            get 
            { 
                return _binding != null && 
                    (_binding.Source != null || _binding.RelativeSource != null || 
                     _binding.ElementName != null || _binding.XPath != null || 
                     _binding.Path != null ); 
            }
        }

        /// <summary>
        /// Is this ResxExtension being used as a multi-binding parent
        /// </summary>
        private bool IsMultiBindingParent
        {
            get { return _children.Count > 0; }
        }

        /// <summary>
        /// Is this ResxExtension being used inside another Resx Extension for multi-binding
        /// </summary>
        private bool IsMultiBindingChild
        {
            get 
            { 
                return (TargetPropertyType == typeof(Collection<ResxExtension>)); 
            }
        }

        /// <summary>
        /// Return the localized resource given a resource Key
        /// </summary>
        /// <param name="resourceKey">The resourceKey</param>
        /// <returns>The value for the current UICulture</returns>
        /// <remarks>Calls GetResource event first then if not handled uses the resource manager</remarks>
        protected virtual object GetLocalizedResource(string resourceKey)
        {
            if (string.IsNullOrEmpty(resourceKey)) return null;
            object result = null;
            if (!string.IsNullOrEmpty(ResxName))
            {
                try
                {
                    if (GetResource != null)
                    {
                        result = GetResource(ResxName, resourceKey, CultureManager.UICulture);
                    }
                    if (result == null)
                    {
                        if (_resourceManager == null)
                        {
                            _resourceManager = GetResourceManager(ResxName);
                        }
                        if (_resourceManager != null)
                        {
                            result = _resourceManager.GetObject(resourceKey, CultureManager.UICulture);
                        }
                    }
                }
                catch
                {
                }
            }
            return result;
        }

         /// <summary>
        /// Return the value for the markup extension
        /// </summary>
        /// <returns>The value from the resources if possible otherwise the default value</returns>
        protected override object GetValue()
        {
            object result = GetLocalizedResource(Key);
            if (result != null && !IsMultiBindingChild)
            {
                try
                {
                    result = ConvertValue(result);
                }
                catch
                {
                }
            }
            if (result == null)
            {
                result = GetDefaultValue(Key);
            }
            return result;
        }

        /// <summary>
        /// Update the given target when the culture changes
        /// </summary>
        /// <param name="target">The target to update</param>
        protected override void UpdateTarget(object target)
        {
            // binding of child extensions is done by the parent
            //
            if (IsMultiBindingChild) return;

            if (IsMultiBindingParent)
            {
                FrameworkElement el = target as FrameworkElement;
                if (el != null)
                {
                    MultiBinding multiBinding = CreateMultiBinding();
                    el.SetBinding(TargetProperty as DependencyProperty, multiBinding);
                }
            }
            else if (IsBindingExpression)
            {
                FrameworkElement el = target as FrameworkElement;
                if (el != null)
                {
                    Binding binding = CreateBinding();
                    el.SetBinding(TargetProperty as DependencyProperty, binding);
                }
            }
            else
            {
                base.UpdateTarget(target);
            }
        }

        /// <summary>
        /// Check if the assembly contains an embedded resx of the given name
        /// </summary>
        /// <param name="assembly">The assembly to check</param>
        /// <param name="resxName">The name of the resource we are looking for</param>
        /// <returns>True if the assembly contains the resource</returns>
        private bool HasEmbeddedResx(Assembly assembly, string resxName)
        {
            // check for dynamic assemblies - we can't call IsDynamic 
            // since it was only introduced in .NET 4
            //
            string assemblyTypeName = assembly.GetType().Name;
            if (assemblyTypeName == "AssemblyBuilder" ||
                assemblyTypeName == "InternalAssemblyBuilder") return false;

            try
            {
                string[] resources = assembly.GetManifestResourceNames();
                string searchName = resxName.ToLower() + ".resources";
                foreach (string resource in resources)
                {
                    if (resource.ToLower() == searchName) return true;
                }
            }
            catch
            {
                // GetManifestResourceNames may throw an exception
                // for some assemblies - just ignore these assemblies.
            }
            return false;
        }

        /// <summary>
        /// Find the assembly that contains the type
        /// </summary>
        /// <returns>The assembly if loaded (otherwise null)</returns>
        private Assembly FindResourceAssembly()
        {
            Assembly assembly = Assembly.GetEntryAssembly();

            // check the entry assembly first - this will short circuit a lot of searching
            //
            if (assembly != null && HasEmbeddedResx(assembly, ResxName)) return assembly;

            var assemblies = AppDomain.CurrentDomain.GetAssemblies();
            foreach (Assembly searchAssembly in assemblies)
            {
                // skip system assemblies
                //
                string name = searchAssembly.FullName;
                if (!name.StartsWith("Microsoft.") &&
                    !name.StartsWith("System.") &&
                    !name.StartsWith("System,") &&
                    !name.StartsWith("mscorlib,") &&
                    !name.StartsWith("PresentationFramework,") &&
                    !name.StartsWith("WindowsBase,"))
                {
                   if (HasEmbeddedResx(searchAssembly, ResxName)) return searchAssembly;
                }
            }
            return null;
        }

        /// <summary>
        /// Get the resource manager for this type
        /// </summary>
        /// <param name="resxName">The name of the embedded resx</param>
        /// <returns>The resource manager</returns>
        /// <remarks>Caches resource managers to improve performance</remarks>
        private ResourceManager GetResourceManager(string resxName)
        {
            WeakReference reference = null;
            ResourceManager result = null;
            if (resxName == null) return null;
            if (_resourceManagers.TryGetValue(resxName, out reference))
            {
                result = reference.Target as ResourceManager;

                // if the resource manager has been garbage collected then remove the cache
                // entry (it will be readded)
                //
                if (result == null)
                {
                    _resourceManagers.Remove(resxName);
                }
            }

            if (result == null)
            {
                Assembly assembly = FindResourceAssembly();
                if (assembly != null)
                {
                    result = new ResourceManager(resxName, assembly);
                }
                _resourceManagers.Add(resxName, new WeakReference(result));
            }
            return result;
        }

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        /// <summary>
        /// Convert a resource object to the type required by the WPF element
        /// </summary>
        /// <param name="value">The resource value to convert</param>
        /// <returns>The WPF element value</returns>
        private object ConvertValue(object value)
        {
            object result = null;
            BitmapSource bitmapSource = null;

            // convert icons and bitmaps to BitmapSource objects that WPF uses
            if (value is Icon)
            {
                Icon icon = value as Icon; 
               
                // For icons we must create a new BitmapFrame from the icon data stream
                // The approach we use for bitmaps (below) doesn't work when setting the
                // Icon property of a window (although it will work for other Icons)
                //
                using (MemoryStream iconStream = new MemoryStream())
                {
                    icon.Save(iconStream);
                    iconStream.Seek(0, SeekOrigin.Begin);
                    bitmapSource = BitmapFrame.Create(iconStream);
                }
            }
            else if (value is Bitmap)
            {
                Bitmap bitmap = value as Bitmap;
                IntPtr bitmapHandle = bitmap.GetHbitmap();
                bitmapSource = Imaging.CreateBitmapSourceFromHBitmap(
                    bitmapHandle, IntPtr.Zero, Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
                bitmapSource.Freeze();
                DeleteObject(bitmapHandle);
            }

            if (bitmapSource != null) 
            {
                // if the target property is expecting the Icon to be content then we
                // create an ImageControl and set its Source property to image
                //
                if (TargetPropertyType == typeof(object))
                {
                    System.Windows.Controls.Image imageControl = new System.Windows.Controls.Image();
                    imageControl.Source = bitmapSource;
                    imageControl.Width = bitmapSource.Width;
                    imageControl.Height = bitmapSource.Height;
                    result = imageControl;
                }
                else
                {
                    result = bitmapSource;
                }
            }
            else
            {
                result = value;
            
                // allow for resources to either contain simple strings or typed data
                //
                Type targetType = TargetPropertyType;
                if (targetType != null)
                {
                    if (value is String && targetType != typeof(String) && targetType != typeof(object))
                    {
                        TypeConverter tc = TypeDescriptor.GetConverter(targetType);
                        result = tc.ConvertFromInvariantString(value as string);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Return the default value for the property
        /// </summary>
        /// <returns></returns>
        private object GetDefaultValue(string key)
        {
            object result = _defaultValue;
            Type targetType = TargetPropertyType;
            if (_defaultValue == null)
            {
                if (targetType == typeof(String) || targetType == typeof(object) || IsMultiBindingChild)
                {
                    result = "#" + key;
                }
            }
            else if (targetType != null)
            {
                // convert the default value if necessary to the required type
                //
                if (targetType != typeof(String) && targetType != typeof(object))
                {
                    try
                    {
                        TypeConverter tc = TypeDescriptor.GetConverter(targetType);
                        result = tc.ConvertFromInvariantString(_defaultValue);
                    }
                    catch
                    {
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Handle a change to the attached DefaultResxName property
        /// </summary>
        /// <param name="element">the dependency object (a WPF element)</param>
        /// <param name="args">the dependency property changed event arguments</param>
        /// <remarks>In design mode update the extension with the correct ResxName</remarks>
        private static void OnDefaultResxNamePropertyChanged(DependencyObject element, DependencyPropertyChangedEventArgs args)
        {
            if (DesignerProperties.GetIsInDesignMode(element))
            {
                foreach (ResxExtension ext in MarkupManager.ActiveExtensions)
                {
                    if (ext.IsTarget(element))
                    {
                        // force the resource manager to be reloaded when the attached resx name changes
                        ext._resourceManager = null;
                        ext._defaultResxName = args.NewValue as string;
                        ext.UpdateTarget(element);
                    }
                }
            }
        }

        #endregion

    }


}
