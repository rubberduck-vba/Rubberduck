﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Rubberduck.Resources.Refactorings {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "16.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    public class EncapsulateField {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal EncapsulateField() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        public static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("Rubberduck.Resources.Refactorings.EncapsulateField", typeof(EncapsulateField).Assembly);
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
        ///   Looks up a localized string similar to Unable to encapsulate &apos;{0}&apos;. ReDim({0}) statement(s) exist in other modules..
        /// </summary>
        public static string ArrayHasExternalRedimFormat {
            get {
                return ResourceManager.GetString("ArrayHasExternalRedimFormat", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Rubberduck - Encapsulate Field.
        /// </summary>
        public static string Caption {
            get {
                return ResourceManager.GetString("Caption", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Select one or more fields to encapsulate.  Accept the default values or edit property names..
        /// </summary>
        public static string InstructionText {
            get {
                return ResourceManager.GetString("InstructionText", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Name conflict.
        /// </summary>
        public static string NameConflictDetected {
            get {
                return ResourceManager.GetString("NameConflictDetected", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Preview:.
        /// </summary>
        public static string Preview {
            get {
                return ResourceManager.GetString("Preview", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to &apos;&lt;===== Property and declaration changes above this line =====&gt;.
        /// </summary>
        public static string PreviewMarker {
            get {
                return ResourceManager.GetString("PreviewMarker", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Creates a property for each UDT Member.
        /// </summary>
        public static string PrivateUDTPropertyText {
            get {
                return ResourceManager.GetString("PrivateUDTPropertyText", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Property Name:.
        /// </summary>
        public static string PropertyName {
            get {
                return ResourceManager.GetString("PropertyName", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Read Only.
        /// </summary>
        public static string ReadOnlyCheckBoxContent {
            get {
                return ResourceManager.GetString("ReadOnlyCheckBoxContent", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Encapsulate Field.
        /// </summary>
        public static string TitleText {
            get {
                return ResourceManager.GetString("TitleText", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Wrap Fields in Private Type.
        /// </summary>
        public static string WrapFieldsInPrivateType {
            get {
                return ResourceManager.GetString("WrapFieldsInPrivateType", resourceCulture);
            }
        }
    }
}
