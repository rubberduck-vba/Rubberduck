using Antlr4.Runtime;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies Functions or Properties referenced by the TruePart(second argument) 
    /// or FalsePart(third argument) of the IIf built-in function.
    /// </summary>
    /// <why>
    /// All arguments of any function/procedure call are always evaluated before the function is invoked so that 
    /// their respective values can be passed as parameters. Even so, the IIf Function's behavior is sometimes mis-interpreted 
    /// to expect that ONLY the 'TruePart' or ONLY the 'FalsePart' expression will be evaluated based on the result of the 
    /// first argument expression. Consequently, the IIf Function can be a source of unanticipated side-effects and errors 
    /// if the user does not account for the fact that both the TruePart and FalsePart arguments are always evaluated.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    ///Private divideByZeroAttempts As Long
    ///
    ///Public Function DoSomeDivision(ByVal dividend As Long, ByVal divisor As Long) As Double
    ///     'IIf will always result in an error with or without valid inputs.
    ///     DoSomeDivision = IIf(divisor > 0, GetQuotient(dividend, divisor), DivideByZeroAttempted())
    ///End Sub
    ///
    ///Private Function GetQuotient(ByVal dividend As Long, ByVal divisor As Long) As Double
    ///     ValidDivision = CDbl(dividend / divisor)
    ///End Function
    ///
    ///Private Function DivideByZeroAttempted() As Double
    ///     DivideByZeroAttempted = 0#
    ///     
    ///     divideByZeroAttempts = divideByZeroAttempts + 1
    ///     
    ///     Err.Raise vbObjectError + 1051, "MyModule", "Divide by Zero attempted"
    ///End Function
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    ///Private isLeapYr As Boolean
    ///
    ///Public Function GetDaysInFebruary(ByVal IsLeapYearFlag As Boolean) As Long
    ///     'Calling IIf results in setting 'isLeapYr' = False regardless of IsLeapYearFlag value as a side-effect of 
    ///     'invoking both the TruePart (LeapYearDays) and the FalsePart (NonLeapYearDays) arguments.
    ///     GetDaysInFebruary = IIf(IsLeapYearFlag, LeapYearDays(), NonLeapYearDays())
    ///End Sub
    ///
    ///Private Function NonLeapYearDays() As Long
    ///     isLeapYr = False
    ///     NonLeapYearDays = 28
    ///End Function
    ///
    ///Private Function LeapYearDays() As Long
    ///     isLeapYr = True
    ///     LeapYearDays = 29
    ///End Function
    ///
    ///Public Property Let IsLeapYear(RHS As Boolean) As Long
    ///     isLeapYr = RHS
    ///End Sub
    ///
    ///Public Property Get IsLeapYear() As Long
    ///     IsLeapYear = isLeapYr
    ///End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    ///Private Const LEAP_YEAR_DAYS As Long = 29
    ///Private Const NON_LEAP_YEAR_DAYS As Long = 28
    ///
    ///Private isLeapYr As Boolean
    ///
    ///Public Function GetDaysInFebruary() As Long
    ///     GetDaysInFebruary = IIf(IsLeapYear, LEAP_YEAR_DAYS, NON_LEAP_YEAR_DAYS)
    ///End Sub
    ///
    ///Public Property Let IsLeapYear(RHS As Boolean) As Long
    ///     isLeapYr = RHS
    ///End Sub
    ///
    ///Public Property Get IsLeapYear() As Long
    ///     IsLeapYear = isLeapYr
    ///End Sub
    /// ]]>
    /// </module>
    /// </example>

    internal sealed class IIfSideEffectInspection : IdentifierReferenceInspectionBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        private static Dictionary<string, string> _nonSideEffectingLibraryFunctionIdentifiers = CreateLibraryFunctionIdentifiersToIgnore();

        public IIfSideEffectInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        //Override DoGetInspectionResults to improve efficiency.  Collates all IIf function TruePart and FalsePart 
        //ArgumentContexts once per module rather than once for every IdentifierReference (of all DeclarationTypes) 
        //contained in a module.
        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var iifReferences = _declarationFinderProvider.DeclarationFinder.BuiltInDeclarations(DeclarationType.Function)
                .SingleOrDefault(d => string.Compare( d.IdentifierName, "IIf", System.StringComparison.InvariantCultureIgnoreCase) == 0)
                .References.Where(rf => rf.QualifiedModuleName == module);

            if (!iifReferences.Any())
            {
                return new List<IInspectionResult>();
            }

            var iifTruePartAndFalsePartArgumentContexts = ExtractTruePartFalsePartArgumentContexts(iifReferences);

            var objectionableReferences = ReferencesInModule(module, finder)
                .Where(reference => reference.Declaration.DeclarationType.HasFlag(DeclarationType.Function)
                    && !_nonSideEffectingLibraryFunctionIdentifiers.ContainsKey(reference.IdentifierName.ToUpperInvariant())
                    && reference.Context.TryGetAncestor<VBAParser.ArgumentContext>(out _)
                    && iifTruePartAndFalsePartArgumentContexts.Any(ac => ac.Contains(reference.Context)));

            return objectionableReferences
                .Select(reference => InspectionResult(reference, finder))
                .ToList();
        }

        //Not used.  This inspection overrides DoGetInspectionResults and aggregates results there.
        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            throw new System.NotImplementedException();
        }

        private static IEnumerable<VBAParser.ArgumentContext> ExtractTruePartFalsePartArgumentContexts(IEnumerable<IdentifierReference> iifReferences)
        {
            (int PositionIndex, string Identifier) truePartParam = (1, "TruePart");
            (int PositionIndex, string Identifier) falsePartParam = (2, "FalsePart");

            var results = new List<VBAParser.ArgumentContext>();

            foreach (var iifReference in iifReferences)
            {
                var argumentContexts = (iifReference.Context.Parent as ParserRuleContext)
                    .GetChild<VBAParser.ArgumentListContext>()
                    .children.OfType<VBAParser.ArgumentContext>()
                    .ToList();

                results.Add(ExtractArgumentContext(argumentContexts, truePartParam));
                results.Add(ExtractArgumentContext(argumentContexts, falsePartParam));
            }

            return results;
        }

        private static VBAParser.ArgumentContext ExtractArgumentContext(IEnumerable<VBAParser.ArgumentContext> argumentContexts, (int PositionIndex, string Identifier) partParam)
        {
            var namedArgumentContexts = argumentContexts.Where(ctxt => ctxt.namedArgument() != null)
                .Select(ctxt => ctxt.namedArgument());

            if (namedArgumentContexts.Any())
            {
                var unrestrictedIDContextsByName = namedArgumentContexts
                    .SelectMany(ctxt => ctxt.children.OfType<VBAParser.UnrestrictedIdentifierContext>())
                    //'ToUpperInvariant' in case the user has (at some point) entered a declaration that re-cased any IIf parameter names
                    .ToDictionary(ch => ch.GetText().ToUpperInvariant());
                
                if (unrestrictedIDContextsByName.TryGetValue(partParam.Identifier.ToUpperInvariant(), out var expressionUnrestrictedIDContext))
                {
                    return expressionUnrestrictedIDContext.Parent.Parent as VBAParser.ArgumentContext;
                }
            }

            return argumentContexts.ElementAt(partParam.PositionIndex);
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            return string.Format(InspectionResults.IIfSideEffectInspection, reference.IdentifierName);
        }

        /// <summary>
        /// Loads VBA Standard library functions that are not be side-effecting or highly unlikely to raise errors
        /// </summary>
        /// <returns></returns>
        private static Dictionary<string, string> CreateLibraryFunctionIdentifiersToIgnore()
        {
            return LoadLibraryFunctionIdentifiersToIgnore( new Dictionary<string, string>(),
                //MS-VBAL 6.1.2.3 Conversion Module
                /*Excluded for potential of raising errors:
                 * "CBool", "CByte", "CCur", "CDate", "CVDate", "CDbl", "CDec", "CInt", "CLng", "CLngLng", "ClngPtr",
                 * "CSng", "CStr", "CVar", "CVErr", "Error","Error$", "Fix", "Hex", "Hex$", "Int", "Oct", "Oct$", "Str", 
                 * "Str$", "Val"
                 */

                //MS-VBAL 6.1.2.4 DateTime Module
                /*Excluded for potential of raising errors: 
                 * "DateAdd", "DateDiff", "DatePart", "DateSerial", "DateValue", "Day", "Hour", "Minute", "Month", "Second", 
                 * "TimeSerial","TimeValue", "Weekday", "Year"
                 */
                "Calendar", "Date", "Date$", "Now", "Time", "Time$", "Timer",
                //MS-VBAL 6.1.2.5 File System
                /*Excluded for potential of raising errors: 
                 * "CurDir", "CurDir$", "Dir", "EOF", "FileAttr", "FileDateTime", "FileLen", "FreeFile", "Loc", "LOF", "Seek" 
                 */

                //MS-VBAL 6.1.2.6 Financial - all excluded
                /*Excluded for potential of raising errors: 
                 * "DDB", "FV", "IPmt", "IRR", "MIRR", "NPer", "NPV", "Pmt", "PPmt", "PV", "Rate", "SLN", "SYD"
                 */

                //MS-VBAL 6.1.2.7 Information
                "IMEStatus", "IsArray", "IsDate", "IsEmpty", "IsError", "IsMissing", "IsNull", "IsNumeric", "IsObject", 
                "QBColor", "RGB", "TypeName", "VarType",

                //MS-VBAL 6.1.2.8 Interaction
                /* Excluded as Potentially side-effecting: 
                 * "CallByName", "Choose", "Command", "Command$", "CreateObject", 
                 * GetObject", "DoEvents", "InputBox", "MsgBox", "Shell", "Switch", "GetAllSettings", "GetAttr", "GetSetting", 
                 * "Partition"
                 */
                "Environ", "Environ$", "IIf",

                //MS-VBAL 6.1.2.10 Math
                /*Excluded for potential of raising errors: 
                 * "Abs", "Atn", "Cos", "Exp", "Log", "Round", "Sgn", "Sin", "Sqr", "Tan"
                */
                "Rnd",

                //MS-VBAL 6.1.2.11 Strings
                /* Excluded for potential of raising errors: 
                 * "Format", "Format$", "FormatDateTime", "FormatNumber", "FormatPercent", "InStr", "InStrB", "InStrRev",
                 * "Join", "Left", "LeftB", "Left$", "LeftB$", "Mid", "MidB", "Mid$", "MidB$", "Replace", "Right", "RightB", 
                 * "Right$", "RightB$", "Asc", "AscW", "AscB", "Chr", "Chr$", "ChB", "ChB$", "ChrW", "ChrW$", "Filter", 
                 * "MonthName", "WeekdayName", "Space", "Space$", "Split","StrConv", "String", "String$"
                */
                "LCase", "LCase$", "Len", "LenB", "Trim", "LTrim", "RTrim", "Trim$", "LTrim$", "RTrim$", "StrComp", 
                "StrReverse", "UCase", "UCase$"
                );
        }

        private static Dictionary<string, string> LoadLibraryFunctionIdentifiersToIgnore(Dictionary<string, string> idMap, params string[] identifiers)
        {
            foreach (var identifier in identifiers)
            {
                idMap.Add(identifier.ToUpperInvariant(), identifier);
            }
            return idMap;
        }
    }
}
