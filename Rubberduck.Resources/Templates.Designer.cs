﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Rubberduck.Resources {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "17.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    public class Templates {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Templates() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        public static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("Rubberduck.Resources.Templates", typeof(Templates).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        public static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Missing Template File.
        /// </summary>
        public static string Menu_Warning_CannotFindTemplate_Caption {
            get {
                return ResourceManager.GetString("Menu_Warning_CannotFindTemplate_Caption", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Cannot find the template for &apos;{0}&apos;. The file may have been renamed or deleted. Expected file name: &apos;{1}&apos;..
        /// </summary>
        public static string Menu_Warning_CannotFindTemplate_Message {
            get {
                return ResourceManager.GetString("Menu_Warning_CannotFindTemplate_Message", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Predeclared Class module (.cls).
        /// </summary>
        public static string PredeclaredClassModule_Caption {
            get {
                return ResourceManager.GetString("PredeclaredClassModule_Caption", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to VERSION 1.0 CLASS
        ///BEGIN
        ///  MultiUse = -1  &apos;True
        ///END
        ///Attribute VB_GlobalNameSpace = False
        ///Attribute VB_Creatable = False
        ///Attribute VB_PredeclaredId = True
        ///Attribute VB_Exposed = False
        ///Attribute VB_Ext_KEY = &quot;Rubberduck&quot;, &quot;Predeclared Class Module&quot;
        ///&apos;@PredeclaredId
        ///Option Explicit.
        /// </summary>
        public static string PredeclaredClassModule_Code {
            get {
                return ResourceManager.GetString("PredeclaredClassModule_Code", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Adds a class module that is predeclared and thus can be used without first creating a new instance..
        /// </summary>
        public static string PredeclaredClassModule_Description {
            get {
                return ResourceManager.GetString("PredeclaredClassModule_Description", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to PredeclaredClassModule.
        /// </summary>
        public static string PredeclaredClassModule_Name {
            get {
                return ResourceManager.GetString("PredeclaredClassModule_Name", resourceCulture);
            }
        }
    }
}
