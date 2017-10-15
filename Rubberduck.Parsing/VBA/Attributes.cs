using System.Collections.Generic;

namespace Rubberduck.Parsing.VBA
{
    //Needed for VB6 support, although 1 and 2 are applicable to VBA.  See 5.2.4.1.1 https://msdn.microsoft.com/en-us/library/ee177292.aspx
    //and 5.2.4.1.2 https://msdn.microsoft.com/en-us/library/ee199159.aspx
    public enum Instancing
    {
        Private = 1,                    //Not exposed via COM.
        PublicNotCreatable = 2,         //TYPEFLAG_FCANCREATE not set.
        SingleUse = 3,                  //TYPEFLAGS.TYPEFLAG_FAPPOBJECT
        GlobalSingleUse = 4,            
        MultiUse = 5,
        GlobalMultiUse = 6
    }

    /// <summary>
    /// A dictionary storing values for a given attribute.
    /// </summary>
    /// <remarks>
    /// Dictionary key is the attribute name/identifier.
    /// </remarks>
    public class Attributes : Dictionary<string, IEnumerable<string>>
    {
        /// <summary>
        /// Specifies that a member is the type's default member.
        /// Only one member in the type is allowed to have this attribute.
        /// </summary>
        /// <param name="identifierName"></param>
        public void AddDefaultMemberAttribute(string identifierName)
        {
            Add(identifierName + ".VB_UserMemId", new[] {"0"});
        }

        public void AddHiddenMemberAttribute(string identifierName)
        {
            Add(identifierName + ".VB_UserMemId", new[] {"40"});
        }

        public void AddEnumeratorMemberAttribute(string identifierName)
        {
            Add(identifierName + ".VB_UserMemId", new[] {"-4"});
        }

        /// <summary>
        /// Corresponds to DISPID_EVALUATE
        /// </summary>
        /// <param name="identifierName"></param>
        public void AddEvaluateMemberAttribute(string identifierName)
        {
            Add(identifierName + ".VB_UserMemId", new[] { "-5" });
        }

        public void AddMemberDescriptionAttribute(string identifierName, string description)
        {
            Add(identifierName + ".VB_Description", new[] {description});
        }

        /// <summary>
        /// Corresponds to TYPEFLAGS.TYPEFLAG_FPREDECLID
        /// </summary>
        public void AddPredeclaredIdTypeAttribute()
        {
            Add("VB_PredeclaredId", new[] {"True"});
        }

        /// <summary>
        /// Corresponds to TYPEFLAG_FAPPOBJECT?
        /// </summary>
        public void AddGlobalClassAttribute()
        {
            Add("VB_GlobalNamespace", new[] {"True"});
        }

        public void AddExposedClassAttribute()
        {
            Add("VB_Exposed", new[] { "True" });
        }

        /// <summary>
        /// Corresponds to *not* having the TYPEFLAG_FNONEXTENSIBLE flag set (which is the default for VBA).
        /// </summary>
        public void AddExtensibledClassAttribute()
        {
            Add("VB_Customizable", new[] { "True" });
        }
    }
}
