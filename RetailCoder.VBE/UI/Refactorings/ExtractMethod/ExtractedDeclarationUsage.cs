using System.Runtime.InteropServices;

namespace Rubberduck.UI.Refactorings.ExtractMethod
{
    /// <summary>
    /// Describes usages of a declared identifier.
    /// </summary>
    public enum ExtractedDeclarationUsage
    {
        /// <summary>
        /// A variable that isn't used in selection, 
        /// will not be extracted.
        /// </summary>
        NotUsed,

        /// <summary>
        /// A variable that is only used in selection, 
        /// will be moved to the extracted method.
        /// </summary>
        UsedOnlyInSelection,
        
        /// <summary>
        /// A variable that is used before selection,
        /// will be extracted as a parameter.
        /// </summary>
        UsedBeforeSelection,
        
        /// <summary>
        /// A variable that is used after selection,
        /// will be extracted as a <c>ByRef</c> parameter 
        /// or become the extracted method's return value.
        /// </summary>
        UsedAfterSelection
    }
}