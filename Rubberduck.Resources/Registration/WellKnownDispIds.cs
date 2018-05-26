namespace Rubberduck.Resources.Registration
{
    // https://msdn.microsoft.com/en-us/library/windows/desktop/ms221242(v=vs.85).aspx
    public static class WellKnownDispIds
    {
        /// <summary>
        /// The Collect property. You use this property if the method you are calling through Invoke is an accessor function.
        /// </summary>
        public const int Collect = -8;

        /// <summary>
        /// The C++ constructor function for the object. 
        /// </summary>
        /// <remarks>
        /// NOTE: Likely used for COM+ only and not applicable for this project.
        /// </remarks>
        public const int Constructor = -6;

        /// <summary>
        /// The C++ destructor function for the object.
        /// </summary>
        /// <remarks>
        /// NOTE: Likely used for COM+ only and not applicable for this project.
        /// </remarks>
        public const int Destructor = -7;

        /// <summary>
        /// The Evaluate method. This method is implicitly invoked when the ActiveX client encloses the arguments in square brackets. 
        /// For example, the following two lines are equivalent:
        /// x.[A1:C1].value = 10
        /// x.Evaluate("A1:C1").value = 10
        /// </summary>
        public const int Evaluate = -5;

        /// <summary>
        /// The _NewEnum property. This special, restricted property is required for collection objects. 
        /// It returns an enumerator object that supports IEnumVARIANT, and should have the restricted attribute specified.
        /// </summary>
        public const int NewEnum = -4;

        /// <summary>
        /// The parameter that receives the value of an assignment in a PROPERTYPUT.
        /// </summary>
        public const int PropertyPut = -3;

        /// <summary>
        /// The value returned by IDispatch::GetIDsOfNames to indicate that a member or parameter name was not found.
        /// </summary>
        /// <remarks>
        /// NOTE: We probably shouldn't use this value for assigning to a member.
        /// </remarks>
        public const int Unknown = -1;

        /// <summary>
        /// The default member for the object. This property or method is invoked when an ActiveX client specifies the 
        /// object name without a property or method. 
        /// </summary>
        public const int Value = 0;
    }
}
