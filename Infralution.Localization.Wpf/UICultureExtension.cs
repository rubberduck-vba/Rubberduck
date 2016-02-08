//
//      FILE:   UICultureExtension.cs.
//
// COPYRIGHT:   Copyright 2008 
//              Infralution
//
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Markup;
using System.Resources;
using System.Reflection;
using System.Windows;
using System.ComponentModel;
using System.Globalization;
using System.Threading;
using System.Diagnostics;
using System.Windows.Interop;
using System.Runtime.InteropServices;

namespace Infralution.Localization.Wpf
{

    /// <summary>
    /// Markup Extension used to dynamically set the Language property of an Markup element to the
    /// the current <see cref="CultureManager.UICulture"/> property value.
    /// </summary>
    /// <remarks>
    /// The culture used for displaying data bound items is based on the Language property.  This
    /// extension allows you to dynamically change the language based on the current 
    /// <see cref="CultureManager.UICulture"/>.
    /// </remarks>
    [MarkupExtensionReturnType(typeof(XmlLanguage))]
    public class UICultureExtension : ManagedMarkupExtension
    {
        /// <summary>
        /// List of active extensions
        /// </summary>
        private static MarkupExtensionManager _markupManager = new MarkupExtensionManager(2);

        /// <summary>
        /// Creates an instance of the extension to set the language property for an
        /// element to the current <see cref="CultureManager.UICulture"/> property value
        /// </summary>
        public UICultureExtension()
            : base(_markupManager)
        {
        }

        /// <summary>
        /// Return the <see cref="XmlLanguage"/> to use for the associated Markup element 
        /// </summary>
        /// <returns>
        /// The <see cref="XmlLanguage"/> corresponding to the current 
        /// <see cref="CultureManager.UICulture"/> property value
        /// </returns>
        protected override object GetValue()
        {
            return XmlLanguage.GetLanguage(CultureManager.UICulture.IetfLanguageTag);
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

    }

}
