//
//      FILE:   ManagedMarkupExtension.cs.
//
// COPYRIGHT:   Copyright 2008 
//              Infralution
//
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Markup;
using System.Reflection;
using System.Windows;
using System.ComponentModel;
using System.Globalization;
using System.Diagnostics;

[assembly: XmlnsDefinition("http://schemas.microsoft.com/winfx/2006/xaml/presentation", "Infralution.Localization.Wpf")]
[assembly: XmlnsDefinition("http://schemas.microsoft.com/winfx/2007/xaml/presentation", "Infralution.Localization.Wpf")]
[assembly: XmlnsDefinition("http://schemas.microsoft.com/winfx/2008/xaml/presentation", "Infralution.Localization.Wpf")]

namespace Infralution.Localization.Wpf
{
    /// <summary>
    /// Defines a base class for markup extensions which are managed by a central 
    /// <see cref="MarkupExtensionManager"/>.   This allows the associated markup targets to be
    /// updated via the manager.
    /// </summary>
    /// <remarks>
    /// The ManagedMarkupExtension holds a weak reference to the target object to allow it to update 
    /// the target.  A weak reference is used to avoid a circular dependency which would prevent the
    /// target being garbage collected.  
    /// </remarks>
    public abstract class ManagedMarkupExtension : MarkupExtension
    {

        #region Member Variables

        /// <summary>
        /// List of weak reference to the target DependencyObjects to allow them to 
        /// be garbage collected
        /// </summary>
        private List<WeakReference> _targetObjects = new List<WeakReference>();

        /// <summary>
        /// The target property 
        /// </summary>
        private object _targetProperty;


        #endregion

        #region Public Interface

        /// <summary>
        /// Create a new instance of the markup extension
        /// </summary>
        public ManagedMarkupExtension(MarkupExtensionManager manager)
        {
            manager.RegisterExtension(this);
        }

        /// <summary>
        /// Return the value for this instance of the Markup Extension
        /// </summary>
        /// <param name="serviceProvider">The service provider</param>
        /// <returns>The value of the element</returns>
        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            RegisterTarget(serviceProvider);
            object result = this;

            // when used in a template the _targetProperty may be null - in this case
            // return this
            //
            if (_targetProperty != null)
            {
                result = GetValue();
            }
            return result;
        }

        /// <summary>
        /// Called by <see cref="ProvideValue(IServiceProvider)"/> to register the target and object
        /// using the extension.   
        /// </summary>
        /// <param name="serviceProvider">The service provider</param>
        protected virtual void RegisterTarget(IServiceProvider serviceProvider)
        {
            var provideValueTarget = serviceProvider.GetService(typeof(IProvideValueTarget)) as IProvideValueTarget;
            object result = this;
            object target = provideValueTarget.TargetObject;

            // Check if the target is a SharedDp which indicates the target is a template
            // In this case we don't register the target and ProvideValue returns this
            // allowing the extension to be evaluated for each instance of the template
            //
            if (target != null && target.GetType().FullName != "System.Windows.SharedDp")
            {
                _targetProperty = provideValueTarget.TargetProperty;
                _targetObjects.Add(new WeakReference(target));
            }
        }

        /// <summary>
        /// Called by <see cref="UpdateTargets"/> to update each target referenced by the extension
        /// </summary>
        /// <param name="target">The target to update</param>
        protected virtual void UpdateTarget(object target)
        {
            if (_targetProperty is DependencyProperty)
            {
                DependencyObject dependencyObject = target as DependencyObject;
                if (dependencyObject != null)
                {
                    dependencyObject.SetValue(_targetProperty as DependencyProperty, GetValue());
                }
            }
            else if (_targetProperty is PropertyInfo)
            {
                (_targetProperty as PropertyInfo).SetValue(target, GetValue(), null);
            }
        }

        /// <summary>
        /// Update the associated targets
        /// </summary>
        public void UpdateTargets()
        {
            foreach (WeakReference reference in _targetObjects)
            {
                if (reference.IsAlive)
                {
                    UpdateTarget(reference.Target);
                }
            }
        }

        /// <summary>
        /// Is the given object the target for the extension 
        /// </summary>
        /// <param name="target">The target to check</param>
        /// <returns>True if the object is one of the targets for this extension</returns>
        public bool IsTarget(object target)
        {
            foreach (WeakReference reference in _targetObjects)
            {
                if (reference.IsAlive && reference.Target == target)
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Is an associated target still alive ie not garbage collected
        /// </summary>
        public bool IsTargetAlive
        {
            get 
            {
                // for normal elements the _targetObjects.Count will always be 1
                // for templates the Count may be zero if this method is called
                // in the middle of window elaboration after the template has been
                // instantiated but before the elements that use it have been.  In
                // this case return true so that we don't unhook the extension
                // prematurely
                //
                if (_targetObjects.Count == 0)
                    return true;
                
                // otherwise just check whether the referenced target(s) are alive
                //
                foreach (WeakReference reference in _targetObjects)
                {
                    if (reference.IsAlive) return true;
                }
                return false; 
            } 
        }

        /// <summary>
        /// Returns true if a target attached to this extension is in design mode
        /// </summary>
        public bool IsInDesignMode
        {
            get
            {
                foreach (WeakReference reference in _targetObjects)
                {
                    DependencyObject element = reference.Target as DependencyObject;
                    if (element != null && DesignerProperties.GetIsInDesignMode(element)) return true;
                }
                return false;
            }
        }

        #endregion

        #region Protected Methods


        /// <summary>
        /// Return the target objects the extension is associated with
        /// </summary>
        /// <remarks>
        /// For normal elements their will be a single target.   For templates
        /// their may be zero or more targets
        /// </remarks>
        protected List<WeakReference> TargetObjects
        {
            get { return _targetObjects; }
        }

        /// <summary>
        /// Return the Target Property the extension is associated with
        /// </summary>
        /// <remarks>
        /// Can either be a <see cref="DependencyProperty"/> or <see cref="PropertyInfo"/>
        /// </remarks>
        protected object TargetProperty
        {
            get { return _targetProperty; }
        }

        /// <summary>
        /// Return the type of the Target Property
        /// </summary>
        protected Type TargetPropertyType
        {
            get
            {
                Type result = null;
                if (_targetProperty is DependencyProperty)
                    result = (_targetProperty as DependencyProperty).PropertyType;
                else if (_targetProperty is PropertyInfo)
                    result = (_targetProperty as PropertyInfo).PropertyType;
                else if (_targetProperty != null)
                    result = _targetProperty.GetType();
                return result;
            }
        }
 
        /// <summary>
        /// Return the value associated with the key from the resource manager
        /// </summary>
        /// <returns>The value from the resources if possible otherwise the default value</returns>
        protected abstract object GetValue();

        #endregion

    }
}
