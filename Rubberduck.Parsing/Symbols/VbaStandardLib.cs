using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Parsing.Symbols
{
    /// <summary>
    /// Defines <see cref="Declaration"/> objects for the standard library.
    /// </summary>
    internal static class VbaStandardLib
    {
        private static IEnumerable<Declaration> _standardLibDeclarations;
        private static readonly QualifiedModuleName VbaModuleName = new QualifiedModuleName("VBA", "VBA");
        private static readonly ICodePaneWrapperFactory WrapperFactory = new CodePaneWrapperFactory();

        public static IEnumerable<Declaration> Declarations
        {
            get
            {
                if (_standardLibDeclarations == null)
                {
                    var nestedTypes = typeof(VbaStandardLib).GetNestedTypes(BindingFlags.NonPublic);
                    var fields = nestedTypes.SelectMany(t => t.GetFields());
                    var values = fields.Select(f => f.GetValue(null));
                    _standardLibDeclarations = values.Cast<Declaration>();
                }

                return _standardLibDeclarations;
            }
        }

        private class VbaLib
        {
            public static readonly Declaration Vba = new Declaration(new QualifiedMemberName(VbaModuleName, "VBA"), null, "VBA", "VBA", true, false, Accessibility.Global, DeclarationType.Project);

            public static readonly Declaration FormShowConstants = new Declaration(new QualifiedMemberName(VbaModuleName, "FormShowConstants"), Vba, "VBA", "FormShowConstants", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbModal = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbModal"), FormShowConstants, "VBA", "FormShowConstants", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbModeless = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbModeless"), FormShowConstants, "VBA", "FormShowConstants", Accessibility.Global, DeclarationType.EnumerationMember, "0");

            public static readonly Declaration VbAppWinStyle = new Declaration(new QualifiedMemberName(VbaModuleName, "VbAppWinStyle"), Vba, "VBA", "VbAppWinStyle", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbHide = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbHide"), VbAppWinStyle, "VBA", "VbAppWinStyle", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbMaximizedFocus = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMaximizedFocus"), VbAppWinStyle, "VBA", "VbAppWinStyle", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbMinimizedFocus = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMinimizedFocus"), VbAppWinStyle, "VBA", "VbAppWinStyle", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbMinimizedNoFocus = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMinimizedNoFocus"), VbAppWinStyle, "VBA", "VbAppWinStyle", Accessibility.Global, DeclarationType.EnumerationMember, "6");
            public static Declaration VbNormalFocus = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbNormalFocus"), VbAppWinStyle, "VBA", "VbAppWinStyle", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbNormalNoFocus = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbNormalNoFocus"), VbAppWinStyle, "VBA", "VbAppWinStyle", Accessibility.Global, DeclarationType.EnumerationMember, "4");

            public static readonly Declaration VbCalendar = new Declaration(new QualifiedMemberName(VbaModuleName, "VbCalendar"), Vba, "VBA", "VbCalendar", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbCalGreg = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbCalGreg"), VbCalendar, "VBA", "VbCalendar", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbCalHijri = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbCalHijri"), VbCalendar, "VBA", "VbCalendar", Accessibility.Global, DeclarationType.EnumerationMember, "1");

            public static readonly Declaration VbCallType = new Declaration(new QualifiedMemberName(VbaModuleName, "VbCallType"), Vba, "VBA", "VbCallType", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbGet = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbGet"), VbCallType, "VBA", "VbCallType", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbLet = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbLet"), VbCallType, "VBA", "VbCallType", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration VbMethod = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMethod"), VbCallType, "VBA", "VbCallType", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbSet = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbSet"), VbCallType, "VBA", "VbCallType", Accessibility.Global, DeclarationType.EnumerationMember, "8");

            public static readonly Declaration VbCompareMethod = new Declaration(new QualifiedMemberName(VbaModuleName, "VbCompareMethod"), Vba, "VBA", "VbCompareMethod", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbBinaryCompare = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbBinaryCompare"), VbCompareMethod, "VBA", "VbCompareMethod", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbTextCompare = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbTextCompare"), VbCompareMethod, "VBA", "VbCompareMethod", Accessibility.Global, DeclarationType.EnumerationMember, "1");

            public static readonly Declaration VbDateTimeFormat = new Declaration(new QualifiedMemberName(VbaModuleName, "VbDateTimeFormat"), Vba, "VBA", "VbDateTimeFormat", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbGeneralDate = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbGeneralDate"), VbDateTimeFormat, "VBA", "VbDateTimeFormat", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbLongDate = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbLongDate"), VbDateTimeFormat, "VBA", "VbDateTimeFormat", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbLongTime = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbLongTime"), VbDateTimeFormat, "VBA", "VbDateTimeFormat", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbShortDate = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbShortDate"), VbDateTimeFormat, "VBA", "VbDateTimeFormat", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbShortTime = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbShortTime"), VbDateTimeFormat, "VBA", "VbDateTimeFormat", Accessibility.Global, DeclarationType.EnumerationMember, "4");

            public static readonly Declaration VbDayOfWeek = new Declaration(new QualifiedMemberName(VbaModuleName, "VbDayOfWeek"), Vba, "VBA", "VbDayOfWeek", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbFriday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFriday"), VbDayOfWeek, "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "6");
            public static Declaration VbMonday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMonday"), VbDayOfWeek, "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbSaturday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbSaturday"), VbDayOfWeek, "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "7");
            public static Declaration VbSunday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbSunday"), VbDayOfWeek, "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbThursday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbThursday"), VbDayOfWeek, "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "5");
            public static Declaration VbTuesday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbTuesday"), VbDayOfWeek, "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbUseSystemDayOfWeek = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbUseSystemDayOfWeek"), VbDayOfWeek, "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbWednesday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbWednesday"), VbDayOfWeek, "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "4");

            public static readonly Declaration VbFileAttribute = new Declaration(new QualifiedMemberName(VbaModuleName, "VbFileAttribute"), Vba, "VBA", "VbFileAttribute", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbNormal = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbNormal"), VbFileAttribute, "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbReadOnly = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbReadOnly"), VbFileAttribute, "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbHidden = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbHidden"), VbFileAttribute, "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbSystem = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbSystem"), VbFileAttribute, "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration VbVolume = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbVolume"), VbFileAttribute, "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "8");
            public static Declaration VbDirectory = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDirectory"), VbFileAttribute, "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "16");
            public static Declaration VbArchive = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbArchive"), VbFileAttribute, "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "32");
            public static Declaration VbAlias = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbAlias"), VbFileAttribute, "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "64");

            public static readonly Declaration VbFirstWeekOfYear = new Declaration(new QualifiedMemberName(VbaModuleName, "VbFirstWeekOfYear"), Vba, "VBA", "VbFirstWeekOfYear", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbFirstFourDays = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFirstFourDays"), VbFirstWeekOfYear, "VBA", "VbFirstWeekOfYear", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbFirstFullWeek = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFirstFullWeek"), VbFirstWeekOfYear, "VBA", "VbFirstWeekOfYear", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbFirstJan1 = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFirstJan1"), VbFirstWeekOfYear, "VBA", "VbFirstWeekOfYear", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbUseSystem = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbUseSystem"), VbFirstWeekOfYear, "VBA", "VbFirstWeekOfYear", Accessibility.Global, DeclarationType.EnumerationMember, "0");

            public static readonly Declaration VbIMEStatus = new Declaration(new QualifiedMemberName(VbaModuleName, "VbIMEStatus"), Vba, "VBA", "VbIMEStatus", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbIMEAlphaDbl = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEAlphaDbl"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "7");
            public static Declaration VbIMEAlphaSng = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEAlphaSng"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "8");
            public static Declaration VbIMEDisable = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEDisable"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbIMEHiragana = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEHiragana"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration VbIMEKatakanaDbl = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEKatakanaDbl"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "5");
            public static Declaration VbIMEKatakanaSng = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEKatakanaSng"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "6");
            public static Declaration VbIMEModeAlpha = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeAlpha"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "8");
            public static Declaration VbIMEModeAlphaFull = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeAlphaFull"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "7");
            public static Declaration VbIMEModeDisable = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeDisable"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbIMEModeHangul = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeHangul"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "10");
            public static Declaration VbIMEModeHangulFull = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeHangulFull"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "9");
            public static Declaration VbIMEModeHiragana = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeHiragana"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration VbIMEModeKatakana = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeKatakana"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "5");
            public static Declaration VbIMEModeKatakanaHalf = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeKatakanaHalf"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "6");
            public static Declaration VbIMEModeNoControl = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeNoControl"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbIMEModeOff = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeOff"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbIMEModeOn = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeOn"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbIMENoOp = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMENoOp"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbIMEOff = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEOff"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbIMEOn = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEOn"), VbIMEStatus, "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "1");

            public static readonly Declaration VbMsgBoxResult = new Declaration(new QualifiedMemberName(VbaModuleName, "VbMsgBoxResult"), Vba, "VBA", "VbMsgBoxResult", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbAbort = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbAbort"), VbMsgBoxResult, "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbCancel = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbCancel"), VbMsgBoxResult, "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbIgnore = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIgnore"), VbMsgBoxResult, "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "5");
            public static Declaration VbNo = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbNo"), VbMsgBoxResult, "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "7");
            public static Declaration VbOk = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbOk"), VbMsgBoxResult, "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbRetry = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbRetry"), VbMsgBoxResult, "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration VbYes = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbYes"), VbMsgBoxResult, "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "6");

            public static readonly Declaration VbMsgBoxStyle = new Declaration(new QualifiedMemberName(VbaModuleName, "VbMsgBoxStyle"), Vba, "VBA", "VbMsgBoxStyle", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbAbortRetryIgnore = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbAbortRetryIgnore"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbApplicationModal = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbApplicationModal"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbCritical = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbCritical"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "16");
            public static Declaration VbDefaultButton1 = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDefaultButton1"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbDefaultButton2 = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDefaultButton2"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "256");
            public static Declaration VbDefaultButton3 = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDefaultButton3"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "512");
            public static Declaration VbDefaultButton4 = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDefaultButton4"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "768");
            public static Declaration VbExclamation = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbExclamation"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "48");
            public static Declaration VbInformation = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbInformation"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "64");
            public static Declaration VbMsgBoxHelpButton = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMsgBoxHelpButton"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "16384");
            public static Declaration VbMsgBoxRight = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMsgBoxRight"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "524288");
            public static Declaration VbMsgBoxRtlReading = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMsgBoxRtlReading"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "1048576");
            public static Declaration VbMsgBoxSetForeground = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMsgBoxSetForeground"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "65536");
            public static Declaration VbOkCancel = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbOkCancel"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbOkOnly = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbOkOnly"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbQuestion = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbQuestion"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "32");
            public static Declaration VbRetryCancel = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbRetryCancel"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "5");
            public static Declaration VbSystemModal = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbSystemModal"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "4096");
            public static Declaration VbYesNo = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbYesNo"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration VbYesNoCancel = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbYesNoCancel"), VbMsgBoxStyle, "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "3");

            public static readonly Declaration VbQueryClose = new Declaration(new QualifiedMemberName(VbaModuleName, "VbQueryClose"), Vba, "VBA", "VbQueryClose", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbAppTaskManager = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbAppTaskManager"), VbQueryClose, "VBA", "VbQueryClose", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbAppWindows = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbAppWindows"), VbQueryClose, "VBA", "VbQueryClose", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbFormCode = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFormCode"), VbQueryClose, "VBA", "VbQueryClose", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbFormControlMenu = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFormControlMenu"), VbQueryClose, "VBA", "VbQueryClose", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbFormMDIForm = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFormMDIForm"), VbQueryClose, "VBA", "VbQueryClose", Accessibility.Global, DeclarationType.EnumerationMember, "4");

            public static readonly Declaration VbStrConv = new Declaration(new QualifiedMemberName(VbaModuleName, "VbStrConv"), Vba, "VBA", "VbStrConv", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbFromUnicode = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFromUnicode"), VbStrConv, "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "128");
            public static Declaration VbHiragana = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbHiragana"), VbStrConv, "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "32");
            public static Declaration VbKatakana = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbKatakana"), VbStrConv, "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "16");
            public static Declaration VbLowerCase = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbLowerCase"), VbStrConv, "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbNarrow = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbNarrow"), VbStrConv, "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "8");
            public static Declaration VbProperCase = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbProperCase"), VbStrConv, "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbUnicode = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbUnicode"), VbStrConv, "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "64");
            public static Declaration VbUpperCase = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbUpperCase"), VbStrConv, "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbWide = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbWide"), VbStrConv, "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "4");

            public static readonly Declaration VbTriState = new Declaration(new QualifiedMemberName(VbaModuleName, "VbTriState"), Vba, "VBA", "VbTriState", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbFalse = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFalse"), VbTriState, "VBA", "VbTriState", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbTrue = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbTrue"), VbTriState, "VBA", "VbTriState", Accessibility.Global, DeclarationType.EnumerationMember, "-1");
            public static Declaration VbUseDefault = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbUseDefault"), VbTriState, "VBA", "VbTriState", Accessibility.Global, DeclarationType.EnumerationMember, "-2");

            public static readonly Declaration VbVarType = new Declaration(new QualifiedMemberName(VbaModuleName, "VbVarType"), Vba, "VBA", "VbVarType", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbArray = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbArray"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "8192");
            public static Declaration VbBoolean = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbBoolean"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "11");
            public static Declaration VbByte = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbByte"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "17");
            public static Declaration VbCurrency = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbCurrency"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "6");
            public static Declaration VbDataObject = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDataObject"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "13");
            public static Declaration VbDate = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDate"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "7");
            public static Declaration VbDecimal = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDecimal"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "14");
            public static Declaration VbDouble = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDouble"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "5");
            public static Declaration VbEmpty = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbEmpty"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbError = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbError"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "10");
            public static Declaration VbInteger = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbInteger"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbLong = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbLong"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbLongLong = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbLongLong"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "20");
            public static Declaration VbNull = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbNull"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbObject = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbObject"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "9");
            public static Declaration VbSingle = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbSingle"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration VbString = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbString"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "8");
            public static Declaration VbUserDefinedType = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbUserDefinedType"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "36");
            public static Declaration VbVariant = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbVariant"), VbVarType, "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "12");
        }

        #region Predefined standard/procedural modules

        private class ColorConstantsModule
        {
            private static readonly QualifiedModuleName ColorConstantsModuleName = new QualifiedModuleName("VBA", "ColorConstants");
            public static readonly Declaration ColorConstants = new Declaration(new QualifiedMemberName(ColorConstantsModuleName, "ColorConstants"), VbaLib.Vba, "VBA", "ColorConstants", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration VbBlack = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbBlack"), ColorConstants, "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "0");
            public static Declaration VbBlue = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbBlue"), ColorConstants, "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "16711680");
            public static Declaration VbCyan = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbCyan"), ColorConstants, "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "16776960");
            public static Declaration VbGreen = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbGreen"), ColorConstants, "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "65280");
            public static Declaration VbMagenta = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbMagenta"), ColorConstants, "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "16711935");
            public static Declaration VbRed = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbRed"), ColorConstants, "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "255");
            public static Declaration VbWhite = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbWhite"), ColorConstants, "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "16777215");
            public static Declaration VbYellow = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbYellow"), ColorConstants, "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "65535");
        }

        private class ConstantsModule
        {
            private static readonly QualifiedModuleName ConstantsModuleName = new QualifiedModuleName("VBA", "Constants");
            public static readonly Declaration Constants = new Declaration(new QualifiedMemberName(ConstantsModuleName, "Constants"), VbaLib.Vba, "VBA", "Constants", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration VbBack = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbBack"), Constants, "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbCr = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbCr"), Constants, "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbCrLf = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbCrLf"), Constants, "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbFormFeed = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbFormFeed"), Constants, "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbLf = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbLf"), Constants, "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbNewLine = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbNewLine"), Constants, "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbNullChar = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbNullChar"), Constants, "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbTab = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbTab"), Constants, "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbVerticalTab = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbVerticalTab"), Constants, "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbNullString = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbNullString"), Constants, "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbObjectError = new ValuedDeclaration(new QualifiedMemberName(ConstantsModuleName, "vbObjectError"), Constants, "VBA.Constants", "Long", Accessibility.Global, DeclarationType.Constant, "-2147221504");
        }

        private class ConversionModule
        {
            private static readonly QualifiedModuleName ConversionModuleName = new QualifiedModuleName("VBA", "Conversion");
            public static readonly Declaration Conversion = new Declaration(new QualifiedMemberName(ConversionModuleName, "Conversion"), VbaLib.Vba, "VBA", "Conversion", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration CBool = new Declaration(new QualifiedMemberName(ConversionModuleName, "CBool"), Conversion, "VBA.Conversion", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CByte = new Declaration(new QualifiedMemberName(ConversionModuleName, "CByte"), Conversion, "VBA.Conversion", "Byte", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CCur = new Declaration(new QualifiedMemberName(ConversionModuleName, "CCur"), Conversion, "VBA.Conversion", "Currency", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CDate = new Declaration(new QualifiedMemberName(ConversionModuleName, "CDate"), Conversion, "VBA.Conversion", "Date", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CVDate = new Declaration(new QualifiedMemberName(ConversionModuleName, "CVDate"), Conversion, "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CDbl = new Declaration(new QualifiedMemberName(ConversionModuleName, "CDbl"), Conversion, "VBA.Conversion", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CDec = new Declaration(new QualifiedMemberName(ConversionModuleName, "CDec"), Conversion, "VBA.Conversion", "Decimal", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CInt = new Declaration(new QualifiedMemberName(ConversionModuleName, "CInt"), Conversion, "VBA.Conversion", "Integer", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CLng = new Declaration(new QualifiedMemberName(ConversionModuleName, "CLng"), Conversion, "VBA.Conversion", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CLngLng = new Declaration(new QualifiedMemberName(ConversionModuleName, "CLngLng"), Conversion, "VBA.Conversion", "LongLong", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CLngPtr = new Declaration(new QualifiedMemberName(ConversionModuleName, "CLngPtr"), Conversion, "VBA.Conversion", "LongPtr", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CSng = new Declaration(new QualifiedMemberName(ConversionModuleName, "CSng"), Conversion, "VBA.Conversion", "Single", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CStr = new Declaration(new QualifiedMemberName(ConversionModuleName, "CStr"), Conversion, "VBA.Conversion", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CVar = new Declaration(new QualifiedMemberName(ConversionModuleName, "CVar"), Conversion, "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CVErr = new Declaration(new QualifiedMemberName(ConversionModuleName, "CVErr"), Conversion, "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Error = new Declaration(new QualifiedMemberName(ConversionModuleName, "Error"), Conversion, "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration ErrorStr = new Declaration(new QualifiedMemberName(ConversionModuleName, "Error$"), Conversion, "VBA.Conversion", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Fix = new Declaration(new QualifiedMemberName(ConversionModuleName, "Fix"), Conversion, "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Hex = new Declaration(new QualifiedMemberName(ConversionModuleName, "Hex"), Conversion, "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration HexStr = new Declaration(new QualifiedMemberName(ConversionModuleName, "Hex$"), Conversion, "VBA.Conversion", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Int = new Declaration(new QualifiedMemberName(ConversionModuleName, "Int"), Conversion, "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration MacID = new Declaration(new QualifiedMemberName(ConversionModuleName, "MacID"), Conversion, "VBA.Conversion", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Oct = new Declaration(new QualifiedMemberName(ConversionModuleName, "Oct"), Conversion, "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration OctStr = new Declaration(new QualifiedMemberName(ConversionModuleName, "Oct$"), Conversion, "VBA.Conversion", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Str = new Declaration(new QualifiedMemberName(ConversionModuleName, "Str"), Conversion, "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration StrStr = new Declaration(new QualifiedMemberName(ConversionModuleName, "Str$"), Conversion, "VBA.Conversion", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Val = new Declaration(new QualifiedMemberName(ConversionModuleName, "Val"), Conversion, "VBA.Conversion", "Double", false, false, Accessibility.Global, DeclarationType.Function);
        }

        private class DateTimeModule
        {
            private static readonly QualifiedModuleName DateTimeModuleName = new QualifiedModuleName("VBA", "DateTime");
            // functions
            public static readonly Declaration DateTime = new Declaration(new QualifiedMemberName(DateTimeModuleName, "DateTime"), VbaLib.Vba, "VBA", "DateTime", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration DateAdd = new Declaration(new QualifiedMemberName(DateTimeModuleName, "DateAdd"), DateTime, "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration DateDiff = new Declaration(new QualifiedMemberName(DateTimeModuleName, "DateDiff"), DateTime, "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration DatePart = new Declaration(new QualifiedMemberName(DateTimeModuleName, "DatePart"), DateTime, "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration DateSerial = new Declaration(new QualifiedMemberName(DateTimeModuleName, "DateSerial"), DateTime, "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration DateValue = new Declaration(new QualifiedMemberName(DateTimeModuleName, "DateValue"), DateTime, "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Day = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Day"), DateTime, "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Hour = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Hour"), DateTime, "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Minute = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Minute"), DateTime, "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Month = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Month"), DateTime, "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Second = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Second"), DateTime, "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration TimeSerial = new Declaration(new QualifiedMemberName(DateTimeModuleName, "TimeSerial"), DateTime, "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration TimeValue = new Declaration(new QualifiedMemberName(DateTimeModuleName, "TimeValue"), DateTime, "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration WeekDay = new Declaration(new QualifiedMemberName(DateTimeModuleName, "WeekDay"), DateTime, "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Year = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Year"), DateTime, "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            // properties
            public static Declaration Calendar = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Calendar"), DateTime, "VBA.DateTime", "vbCalendar", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Date = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Date"), DateTime, "VBA.DateTime", "Date", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration DateStr = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Date$"), DateTime, "VBA.DateTime", "String", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Now = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Now"), DateTime, "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Time = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Time"), DateTime, "VBA.DateTime", "Date", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration TimeStr = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Time$"), DateTime, "VBA.DateTime", "String", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Timer = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Timer"), DateTime, "VBA.DateTime", "Single", false, false, Accessibility.Global, DeclarationType.PropertyGet);
        }

        private class FileSystemModule
        {
            private static readonly QualifiedModuleName FileSystemModuleName = new QualifiedModuleName("VBA", "FileSystem");
            // functions
            public static readonly Declaration FileSystem = new Declaration(new QualifiedMemberName(FileSystemModuleName, "FileSystem"), VbaLib.Vba, "VBA", "FileSystem", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration CurDir = new Declaration(new QualifiedMemberName(FileSystemModuleName, "CurDir"), FileSystem, "VBA.FileSystem", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CurDirStr = new Declaration(new QualifiedMemberName(FileSystemModuleName, "CurDir$"), FileSystem, "VBA.FileSystem", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Dir = new Declaration(new QualifiedMemberName(FileSystemModuleName, "Dir"), FileSystem, "VBA.FileSystem", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration EOF = new Declaration(new QualifiedMemberName(FileSystemModuleName, "EOF"), FileSystem, "VBA.FileSystem", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FileAttr = new Declaration(new QualifiedMemberName(FileSystemModuleName, "FileAttr"), FileSystem, "VBA.FileSystem", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FileDateTime = new Declaration(new QualifiedMemberName(FileSystemModuleName, "FileDateTime"), FileSystem, "VBA.FileSystem", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FileLen = new Declaration(new QualifiedMemberName(FileSystemModuleName, "FileLen"), FileSystem, "VBA.FileSystem", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FreeFile = new Declaration(new QualifiedMemberName(FileSystemModuleName, "FreeFile"), FileSystem, "VBA.FileSystem", "Integer", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Loc = new Declaration(new QualifiedMemberName(FileSystemModuleName, "Loc"), FileSystem, "VBA.FileSystem", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration LOF = new Declaration(new QualifiedMemberName(FileSystemModuleName, "LOF"), FileSystem, "VBA.FileSystem", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Seek = new Declaration(new QualifiedMemberName(FileSystemModuleName, "Seek"), FileSystem, "VBA.FileSystem", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            // procedures
            public static Declaration ChDir = new Declaration(new QualifiedMemberName(FileSystemModuleName, "ChDir"), FileSystem, "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration ChDrive = new Declaration(new QualifiedMemberName(FileSystemModuleName, "ChDrive"), FileSystem, "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration FileCopy = new Declaration(new QualifiedMemberName(FileSystemModuleName, "FileCopy"), FileSystem, "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration Kill = new Declaration(new QualifiedMemberName(FileSystemModuleName, "Kill"), FileSystem, "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration MkDir = new Declaration(new QualifiedMemberName(FileSystemModuleName, "MkDir"), FileSystem, "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration RmDir = new Declaration(new QualifiedMemberName(FileSystemModuleName, "RmDir"), FileSystem, "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration SetAttr = new Declaration(new QualifiedMemberName(FileSystemModuleName, "SetAttr"), FileSystem, "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure);
        }

        private class FinancialModule
        {
            private static readonly QualifiedModuleName FinancialModuleName = new QualifiedModuleName("VBA", "Financial");
            public static readonly Declaration Financial = new Declaration(new QualifiedMemberName(FinancialModuleName, "Financial"), VbaLib.Vba, "VBA", "Financial", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration DDB = new Declaration(new QualifiedMemberName(FinancialModuleName, "DDB"), Financial, "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FV = new Declaration(new QualifiedMemberName(FinancialModuleName, "FV"), Financial, "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IPmt = new Declaration(new QualifiedMemberName(FinancialModuleName, "IPmt"), Financial, "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IRR = new Declaration(new QualifiedMemberName(FinancialModuleName, "IRR"), Financial, "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration MIRR = new Declaration(new QualifiedMemberName(FinancialModuleName, "MIRR"), Financial, "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration NPer = new Declaration(new QualifiedMemberName(FinancialModuleName, "NPer"), Financial, "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration NPV = new Declaration(new QualifiedMemberName(FinancialModuleName, "NPV"), Financial, "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Pmt = new Declaration(new QualifiedMemberName(FinancialModuleName, "Pmt"), Financial, "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration PPmt = new Declaration(new QualifiedMemberName(FinancialModuleName, "PPmt"), Financial, "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration PV = new Declaration(new QualifiedMemberName(FinancialModuleName, "PV"), Financial, "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Rate = new Declaration(new QualifiedMemberName(FinancialModuleName, "Rate"), Financial, "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration SLN = new Declaration(new QualifiedMemberName(FinancialModuleName, "SLN"), Financial, "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration SYD = new Declaration(new QualifiedMemberName(FinancialModuleName, "SYD"), Financial, "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
        }

        private class HiddenModule
        {
            private static readonly QualifiedModuleName HiddenModuleName = new QualifiedModuleName("VBA", "[_HiddenModule]");
            public static readonly Declaration Hidden = new Declaration(new QualifiedMemberName(HiddenModuleName, "[_HiddenModule]"), VbaLib.Vba, "VBA", "[_HiddenModule]", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration Array = new Declaration(new QualifiedMemberName(HiddenModuleName, "Array"), Hidden, "VBA.[_HiddenModule]", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Input = new Declaration(new QualifiedMemberName(HiddenModuleName, "Input"), Hidden, "VBA.[_HiddenModule]", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration InputStr = new Declaration(new QualifiedMemberName(HiddenModuleName, "Input$"), Hidden, "VBA.[_HiddenModule]", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration InputB = new Declaration(new QualifiedMemberName(HiddenModuleName, "InputB"), Hidden, "VBA.[_HiddenModule]", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration InputBStr = new Declaration(new QualifiedMemberName(HiddenModuleName, "InputB$"), Hidden, "VBA.[_HiddenModule]", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Width = new Declaration(new QualifiedMemberName(HiddenModuleName, "Width"), Hidden, "VBA.[_HiddenModule]", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            
            // hidden members... of hidden module (like, very very hidden!)
            public static Declaration ObjPtr = new Declaration(new QualifiedMemberName(HiddenModuleName, "ObjPtr"), Hidden, "VBA.[_HiddenModule]", "LongPtr", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration StrPtr = new Declaration(new QualifiedMemberName(HiddenModuleName, "StrPtr"), Hidden, "VBA.[_HiddenModule]", "LongPtr", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration VarPtr = new Declaration(new QualifiedMemberName(HiddenModuleName, "VarPtr"), Hidden, "VBA.[_HiddenModule]", "LongPtr", false, false, Accessibility.Global, DeclarationType.Function);
        }

        private class InformationModule
        {
            private static readonly QualifiedModuleName InformationModuleName = new QualifiedModuleName("VBA", "Information");
            public static readonly Declaration Information = new Declaration(new QualifiedMemberName(InformationModuleName, "Information"), VbaLib.Vba, "VBA", "Information", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration Err = new Declaration(new QualifiedMemberName(InformationModuleName, "Err"), Information, "VBA.Information", "ErrObject", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Erl = new Declaration(new QualifiedMemberName(InformationModuleName, "Erl"), Information, "VBA.Information", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IMEStatus = new Declaration(new QualifiedMemberName(InformationModuleName, "IMEStatus"), Information, "VBA.Information", "vbIMEStatus", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IsArray = new Declaration(new QualifiedMemberName(InformationModuleName, "IsArray"), Information, "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IsDate = new Declaration(new QualifiedMemberName(InformationModuleName, "IsDate"), Information, "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IsEmpty = new Declaration(new QualifiedMemberName(InformationModuleName, "IsEmpty"), Information, "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IsError = new Declaration(new QualifiedMemberName(InformationModuleName, "IsError"), Information, "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IsMissing = new Declaration(new QualifiedMemberName(InformationModuleName, "IsMissing"), Information, "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IsNull = new Declaration(new QualifiedMemberName(InformationModuleName, "IsNull"), Information, "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IsNumeric = new Declaration(new QualifiedMemberName(InformationModuleName, "IsNumeric"), Information, "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IsObject = new Declaration(new QualifiedMemberName(InformationModuleName, "IsObject"), Information, "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration QBColor = new Declaration(new QualifiedMemberName(InformationModuleName, "QBColor"), Information, "VBA.Information", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration RGB = new Declaration(new QualifiedMemberName(InformationModuleName, "RGB"), Information, "VBA.Information", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration TypeName = new Declaration(new QualifiedMemberName(InformationModuleName, "TypeName"), Information, "VBA.Information", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration VarType = new Declaration(new QualifiedMemberName(InformationModuleName, "VarType"), Information, "VBA.Information", "vbVarType", false, false, Accessibility.Global, DeclarationType.Function);
        }

        private class InteractionModule
        {
            private static readonly QualifiedModuleName InteractionModuleName = new QualifiedModuleName("VBA", "Interaction");
            // functions
            public static readonly Declaration Interaction = new Declaration(new QualifiedMemberName(InteractionModuleName, "Interaction"), VbaLib.Vba, "VBA", "Interaction", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration CallByName = new Declaration(new QualifiedMemberName(InteractionModuleName, "CallByName"), Interaction, "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Choose = new Declaration(new QualifiedMemberName(InteractionModuleName, "Choose"), Interaction, "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Command = new Declaration(new QualifiedMemberName(InteractionModuleName, "Command"), Interaction, "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CommandStr = new Declaration(new QualifiedMemberName(InteractionModuleName, "Command$"), Interaction, "VBA.Interaction", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CreateObject = new Declaration(new QualifiedMemberName(InteractionModuleName, "CreateObject"), Interaction, "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration DoEvents = new Declaration(new QualifiedMemberName(InteractionModuleName, "DoEvents"), Interaction, "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Environ = new Declaration(new QualifiedMemberName(InteractionModuleName, "Environ"), Interaction, "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration EnvironStr = new Declaration(new QualifiedMemberName(InteractionModuleName, "Environ$"), Interaction, "VBA.Interaction", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration GetAllSettings = new Declaration(new QualifiedMemberName(InteractionModuleName, "GetAllSettings"), Interaction, "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration GetAttr = new Declaration(new QualifiedMemberName(InteractionModuleName, "GetAttr"), Interaction, "VBA.Interaction", "vbFileAttribute", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration GetObject = new Declaration(new QualifiedMemberName(InteractionModuleName, "GetObject"), Interaction, "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration GetSetting = new Declaration(new QualifiedMemberName(InteractionModuleName, "GetSetting"), Interaction, "VBA.Interaction", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IIf = new Declaration(new QualifiedMemberName(InteractionModuleName, "IIf"), Interaction, "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration InputBox = new Declaration(new QualifiedMemberName(InteractionModuleName, "InputBox"), Interaction, "VBA.Interaction", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration MacScript = new Declaration(new QualifiedMemberName(InteractionModuleName, "MacScript"), Interaction, "VBA.Interaction", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration MsgBox = new Declaration(new QualifiedMemberName(InteractionModuleName, "MsgBox"), Interaction, "VBA.Interaction", "vbMsgBoxResult", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Partition = new Declaration(new QualifiedMemberName(InteractionModuleName, "Partition"), Interaction, "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Shell = new Declaration(new QualifiedMemberName(InteractionModuleName, "Shell"), Interaction, "VBA.Interaction", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Switch = new Declaration(new QualifiedMemberName(InteractionModuleName, "Switch"), Interaction, "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            // procedures
            public static Declaration AppActivate = new Declaration(new QualifiedMemberName(InteractionModuleName, "AppActivate"), Interaction, "VBA.Interaction", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration Beep = new Declaration(new QualifiedMemberName(InteractionModuleName, "Beep"), Interaction, "VBA.Interaction", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration DeleteSetting = new Declaration(new QualifiedMemberName(InteractionModuleName, "DeleteSetting"), Interaction, "VBA.Interaction", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration SaveSetting = new Declaration(new QualifiedMemberName(InteractionModuleName, "SaveSetting"), Interaction, "VBA.Interaction", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration SendKeys = new Declaration(new QualifiedMemberName(InteractionModuleName, "SendKeys"), Interaction, "VBA.Interaction", null, false, false, Accessibility.Global, DeclarationType.Procedure);
        }

        private class KeyCodeConstantsModule
        {
            private static readonly QualifiedModuleName KeyCodeConstantsModuleName = new QualifiedModuleName("VBA", "KeyCodeConstants");
            public static readonly Declaration KeyCodeConstants = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "KeyCodeConstants"), VbaLib.Vba, "VBA", "KeyCodeConstants", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration VbKeyLButton = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyLButton"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "1");
            public static Declaration VbKeyRButton = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyRButton"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "2");
            public static Declaration VbKeyCancel = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyCancel"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "3");
            public static Declaration VbKeyMButton = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyMButton"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "4");
            public static Declaration VbKeyBack = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyBack"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "8");
            public static Declaration VbKeyTab = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyTab"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "9");
            public static Declaration VbKeyClear = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyClear"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "12");
            public static Declaration VbKeyReturn = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyReturn"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "13");
            public static Declaration VbKeyShift = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyShift"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "16");
            public static Declaration VbKeyControl = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyControl"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "17");
            public static Declaration VbKeyMenu = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyMenu"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "18");
            public static Declaration VbKeyPause = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyPause"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "19");
            public static Declaration VbKeyCapital = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyCapital"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "20");
            public static Declaration VbKeyEscape = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyEscape"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "27");
            public static Declaration VbKeySpace = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeySpace"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "32");
            public static Declaration VbKeyPageUp = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyPageUp"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "33");
            public static Declaration VbKeyPageDown = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyPageDown"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "34");
            public static Declaration VbKeyEnd = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyEnd"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "35");
            public static Declaration VbKeyHome = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyHome"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "36");
            public static Declaration VbKeyLeft = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyLeft"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "37");
            public static Declaration VbKeyUp = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyUp"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "38");
            public static Declaration VbKeyRight = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyRight"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "39");
            public static Declaration VbKeyDown = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyDown"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "40");
            public static Declaration VbKeySelect = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeySelect"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "41");
            public static Declaration VbKeyPrint = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyPrint"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "42");
            public static Declaration VbKeyExecute = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyExecute"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "43");
            public static Declaration VbKeySnapshot = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeySnapshot"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "44");
            public static Declaration VbKeyInsert = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyInsert"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "45");
            public static Declaration VbKeyDelete = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyDelete"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "46");
            public static Declaration VbKeyHelp = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyHelp"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "47");
            public static Declaration VbKeyNumLock = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumLock"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "144");
            public static Declaration VbKeyA = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyA"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "65");
            public static Declaration VbKeyB = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyB"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "66");
            public static Declaration VbKeyC = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyC"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "67");
            public static Declaration VbKeyD = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyD"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "68");
            public static Declaration VbKeyE = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyE"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "69");
            public static Declaration VbKeyF = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "70");
            public static Declaration VbKeyG = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyG"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "71");
            public static Declaration VbKeyH = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyH"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "72");
            public static Declaration VbKeyI = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyI"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "73");
            public static Declaration VbKeyJ = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyJ"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "74");
            public static Declaration VbKeyK = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyK"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "75");
            public static Declaration VbKeyL = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyL"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "76");
            public static Declaration VbKeyM = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyM"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "77");
            public static Declaration VbKeyN = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyN"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "78");
            public static Declaration VbKeyO = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyO"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "79");
            public static Declaration VbKeyP = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyP"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "80");
            public static Declaration VbKeyQ = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyQ"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "81");
            public static Declaration VbKeyR = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyR"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "82");
            public static Declaration VbKeyS = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyS"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "83");
            public static Declaration VbKeyT = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyT"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "84");
            public static Declaration VbKeyU = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyU"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "85");
            public static Declaration VbKeyV = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyV"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "86");
            public static Declaration VbKeyW = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyW"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "87");
            public static Declaration VbKeyX = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyX"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "88");
            public static Declaration VbKeyY = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyY"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "89");
            public static Declaration VbKeyZ = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyZ"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "90");
            public static Declaration VbKey0 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey0"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "48");
            public static Declaration VbKey1 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey1"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "49");
            public static Declaration VbKey2 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey2"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "50");
            public static Declaration VbKey3 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey3"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "51");
            public static Declaration VbKey4 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey4"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "52");
            public static Declaration VbKey5 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey5"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "53");
            public static Declaration VbKey6 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey6"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "54");
            public static Declaration VbKey7 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey7"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "55");
            public static Declaration VbKey8 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey8"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "56");
            public static Declaration VbKey9 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey9"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "57");
            public static Declaration VbKeyNumpad0 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad0"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "96");
            public static Declaration VbKeyNumpad1 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad1"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "97");
            public static Declaration VbKeyNumpad2 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad2"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "98");
            public static Declaration VbKeyNumpad3 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad3"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "99");
            public static Declaration VbKeyNumpad4 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad4"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "100");
            public static Declaration VbKeyNumpad5 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad5"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "101");
            public static Declaration VbKeyNumpad6 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad6"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "102");
            public static Declaration VbKeyNumpad7 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad7"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "103");
            public static Declaration VbKeyNumpad8 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad8"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "104");
            public static Declaration VbKeyNumpad9 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad9"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "105");
            public static Declaration VbKeyMultiply = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyMultiply"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "106");
            public static Declaration VbKeyAdd = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyAdd"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "107");
            public static Declaration VbKeySeparator = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeySeparator"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "108");
            public static Declaration VbKeySubtract = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeySubtract"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "109");
            public static Declaration VbKeyDecimal = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyDecimal"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "110");
            public static Declaration VbKeyDivide = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyDivide"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "111");
            public static Declaration VbKeyF1 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF1"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "112");
            public static Declaration VbKeyF2 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF2"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "113");
            public static Declaration VbKeyF3 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF3"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "114");
            public static Declaration VbKeyF4 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF4"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "115");
            public static Declaration VbKeyF5 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF5"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "116");
            public static Declaration VbKeyF6 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF6"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "117");
            public static Declaration VbKeyF7 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF7"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "118");
            public static Declaration VbKeyF8 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF8"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "119");
            public static Declaration VbKeyF9 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF9"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "120");
            public static Declaration VbKeyF10 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF10"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "121");
            public static Declaration VbKeyF11 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF11"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "122");
            public static Declaration VbKeyF12 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF12"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "123");
            public static Declaration VbKeyF13 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF13"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "124");
            public static Declaration VbKeyF14 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF14"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "125");
            public static Declaration VbKeyF15 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF15"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "126");
            public static Declaration VbKeyF16 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF16"), KeyCodeConstants, "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "127");
        }

        private class MathModule
        {
            private static readonly QualifiedModuleName MathModuleName = new QualifiedModuleName("VBA", "Math");
            // functions
            public static readonly Declaration Math = new Declaration(new QualifiedMemberName(MathModuleName, "Math"), VbaLib.Vba, "VBA", "Math", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration Abs = new Declaration(new QualifiedMemberName(MathModuleName, "Abs"), Math, "VBA.Math", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Atn = new Declaration(new QualifiedMemberName(MathModuleName, "Atn"), Math, "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Cos = new Declaration(new QualifiedMemberName(MathModuleName, "Cos"), Math, "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Exp = new Declaration(new QualifiedMemberName(MathModuleName, "Exp"), Math, "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Log = new Declaration(new QualifiedMemberName(MathModuleName, "Log"), Math, "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Rnd = new Declaration(new QualifiedMemberName(MathModuleName, "Rnd"), Math, "VBA.Math", "Single", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Round = new Declaration(new QualifiedMemberName(MathModuleName, "Round"), Math, "VBA.Math", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Sgn = new Declaration(new QualifiedMemberName(MathModuleName, "Sgn"), Math, "VBA.Math", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Sin = new Declaration(new QualifiedMemberName(MathModuleName, "Sin"), Math, "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Sqr = new Declaration(new QualifiedMemberName(MathModuleName, "Sqr"), Math, "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Tan = new Declaration(new QualifiedMemberName(MathModuleName, "Tan"), Math, "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            //procedures
            public static Declaration Randomize = new Declaration(new QualifiedMemberName(MathModuleName, "Randomize"), Math, "VBA.Math", null, false, false, Accessibility.Global, DeclarationType.Procedure);
        }

        private class StringsModule
        {
            private static readonly QualifiedModuleName StringsModuleName = new QualifiedModuleName("VBA", "Strings");
            public static readonly Declaration Strings = new Declaration(new QualifiedMemberName(StringsModuleName, "Strings"), VbaLib.Vba, "VBA", "Strings", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration Asc = new Declaration(new QualifiedMemberName(StringsModuleName, "Asc"), Strings, "VBA.Strings", "Integer", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration AscW = new Declaration(new QualifiedMemberName(StringsModuleName, "AscW"), Strings, "VBA.Strings", "Integer", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration AscB = new Declaration(new QualifiedMemberName(StringsModuleName, "AscB"), Strings, "VBA.Strings", "Integer", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Chr = new Declaration(new QualifiedMemberName(StringsModuleName, "Chr"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration ChrStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Chr$"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration ChrB = new Declaration(new QualifiedMemberName(StringsModuleName, "ChrB"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration ChrBStr = new Declaration(new QualifiedMemberName(StringsModuleName, "ChrB$"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration ChrW = new Declaration(new QualifiedMemberName(StringsModuleName, "ChrW"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration ChrWStr = new Declaration(new QualifiedMemberName(StringsModuleName, "ChrW$"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Filter = new Declaration(new QualifiedMemberName(StringsModuleName, "Filter"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Format = new Declaration(new QualifiedMemberName(StringsModuleName, "Format"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FormatStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Format$"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FormatCurrency = new Declaration(new QualifiedMemberName(StringsModuleName, "FormatCurrency"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FormatDateTime = new Declaration(new QualifiedMemberName(StringsModuleName, "FormatDateTime"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FormatNumber = new Declaration(new QualifiedMemberName(StringsModuleName, "FormatNumber"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FormatPercent = new Declaration(new QualifiedMemberName(StringsModuleName, "FormatPercent"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration InStr = new Declaration(new QualifiedMemberName(StringsModuleName, "InStr"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration InStrB = new Declaration(new QualifiedMemberName(StringsModuleName, "InStrB"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration InStrRev = new Declaration(new QualifiedMemberName(StringsModuleName, "InStrRev"), Strings, "VBA.Strings", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Join = new Declaration(new QualifiedMemberName(StringsModuleName, "Join"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration LCase = new Declaration(new QualifiedMemberName(StringsModuleName, "LCase"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration LCaseStr = new Declaration(new QualifiedMemberName(StringsModuleName, "LCase$"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Left = new Declaration(new QualifiedMemberName(StringsModuleName, "Left"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration LeftB = new Declaration(new QualifiedMemberName(StringsModuleName, "LeftB"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration LeftStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Left$"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration LeftBStr = new Declaration(new QualifiedMemberName(StringsModuleName, "LeftB$"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Len = new Declaration(new QualifiedMemberName(StringsModuleName, "Len"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration LenB = new Declaration(new QualifiedMemberName(StringsModuleName, "LenB"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration LTrim = new Declaration(new QualifiedMemberName(StringsModuleName, "LTrim"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration RTrim = new Declaration(new QualifiedMemberName(StringsModuleName, "RTrim"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Trim = new Declaration(new QualifiedMemberName(StringsModuleName, "Trim"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration LTrimStr = new Declaration(new QualifiedMemberName(StringsModuleName, "LTrim$"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration RTrimStr = new Declaration(new QualifiedMemberName(StringsModuleName, "RTrim$"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration TrimStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Trim$"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Mid = new Declaration(new QualifiedMemberName(StringsModuleName, "Mid"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration MidB = new Declaration(new QualifiedMemberName(StringsModuleName, "MidB"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration MidStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Mid$"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration MidBStr = new Declaration(new QualifiedMemberName(StringsModuleName, "MidB$"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration MonthName = new Declaration(new QualifiedMemberName(StringsModuleName, "MonthName"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Replace = new Declaration(new QualifiedMemberName(StringsModuleName, "Replace"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Right = new Declaration(new QualifiedMemberName(StringsModuleName, "Right"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration RightB = new Declaration(new QualifiedMemberName(StringsModuleName, "RightB"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration RightStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Right$"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration RightBStr = new Declaration(new QualifiedMemberName(StringsModuleName, "RightB$"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Space = new Declaration(new QualifiedMemberName(StringsModuleName, "Space"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration SpaceStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Space$"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Split = new Declaration(new QualifiedMemberName(StringsModuleName, "Split"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration StrComp = new Declaration(new QualifiedMemberName(StringsModuleName, "StrComp"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration StrConv = new Declaration(new QualifiedMemberName(StringsModuleName, "StrConv"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration String = new Declaration(new QualifiedMemberName(StringsModuleName, "String"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration StringStr = new Declaration(new QualifiedMemberName(StringsModuleName, "String$"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration StrReverse = new Declaration(new QualifiedMemberName(StringsModuleName, "StrReverse"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration UCase = new Declaration(new QualifiedMemberName(StringsModuleName, "UCase"), Strings, "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration UCaseStr = new Declaration(new QualifiedMemberName(StringsModuleName, "UCase$"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration WeekdayName = new Declaration(new QualifiedMemberName(StringsModuleName, "WeekdayName"), Strings, "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
        }

        private class SystemColorConstantsModule
        {
            private static readonly QualifiedModuleName SystemColorConstantsModuleName = new QualifiedModuleName("VBA", "SystemColorConstants");
            public static readonly Declaration SystemColorConstants = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "SystemColorConstants"), VbaLib.Vba, "VBA", "SystemColorConstants", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration VbScrollBars = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbScrollBars"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbDesktop = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbDesktop"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbActiveTitleBar = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbActiveTitleBar"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbInactiveTitleBar = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbInactiveTitleBar"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbMenuBar = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbMenuBar"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbWindowBackground = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbWindowBackground"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbWindowFrame = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbWindowFrame"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbMenuText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbMenuText"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbWindowText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbWindowText"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbTitleBarText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbTitleBarText"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbActiveBorder = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbActiveBorder"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbInactiveBorder = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbInactiveBorder"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbApplicationWorkspace = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbApplicationWorkspace"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbHighlight = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbHighlight"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbHighlightText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbHighlightText"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbButtonFace = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbButtonFace"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbButtonShadow = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbButtonShadow"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbGrayText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbGrayText"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbButtonText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbButtonText"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbInactiveCaptionText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbInactiveCaptionText"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration Vb3DHighlight = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vb3DHighlight"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration Vb3DDKShadow = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vb3DDKShadow"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration Vb3DLight = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vb3DLight"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration Vb3DFace = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vb3DFace"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration Vb3DShadow = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vb3DShadow"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbInfoText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbInfoText"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbInfoBackground = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbInfoBackground"), SystemColorConstants, "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
        }

        #endregion

        #region Predefined class modules

        private class CollectionClass
        {
            public static readonly Declaration Collection = new Declaration(new QualifiedMemberName(VbaModuleName, "Collection"), VbaLib.Vba, "VBA", "Collection", false, false, Accessibility.Global, DeclarationType.Class);
            public static Declaration NewEnum = new Declaration(new QualifiedMemberName(VbaModuleName, "[_NewEnum]"), Collection, "VBA.Collection", "Unknown", false, false, Accessibility.Public, DeclarationType.Function);
            public static Declaration Count = new Declaration(new QualifiedMemberName(VbaModuleName, "Count"), Collection, "VBA.Collection", "Long", false, false, Accessibility.Public, DeclarationType.Function);
            public static Declaration Item = new Declaration(new QualifiedMemberName(VbaModuleName, "Item"), Collection, "VBA.Collection", "Variant", false, false, Accessibility.Public, DeclarationType.Function);
            public static Declaration Add = new Declaration(new QualifiedMemberName(VbaModuleName, "Add"), Collection, "VBA.Collection", null, false, false, Accessibility.Public, DeclarationType.Procedure);
            public static Declaration Remove = new Declaration(new QualifiedMemberName(VbaModuleName, "Remove"), Collection, "VBA.Collection", null, false, false, Accessibility.Public, DeclarationType.Procedure);
        }

        private class ErrObjectClass
        {
            public static readonly Declaration ErrObject = new Declaration(new QualifiedMemberName(VbaModuleName, "ErrObject"), VbaLib.Vba, "VBA", "ErrObject", false, false, Accessibility.Global, DeclarationType.Class);
            public static Declaration Clear = new Declaration(new QualifiedMemberName(VbaModuleName, "Clear"), ErrObject, "VBA.ErrObject", null, false, false, Accessibility.Public, DeclarationType.Procedure);
            public static Declaration Raise = new Declaration(new QualifiedMemberName(VbaModuleName, "Raise"), ErrObject, "VBA.ErrObject", null, false, false, Accessibility.Public, DeclarationType.Procedure);
            public static Declaration Description = new Declaration(new QualifiedMemberName(VbaModuleName, "Description"), ErrObject, "VBA.ErrObject", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration HelpContext = new Declaration(new QualifiedMemberName(VbaModuleName, "HelpContext"), ErrObject, "VBA.ErrObject", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration HelpFile = new Declaration(new QualifiedMemberName(VbaModuleName, "HelpFile"), ErrObject, "VBA.ErrObject", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration LastDllError = new Declaration(new QualifiedMemberName(VbaModuleName, "LastDllError"), ErrObject, "VBA.ErrObject", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration Number = new Declaration(new QualifiedMemberName(VbaModuleName, "Number"), ErrObject, "VBA.ErrObject", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration Source = new Declaration(new QualifiedMemberName(VbaModuleName, "Source"), ErrObject, "VBA.ErrObject", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet);
        }

        private class GlobalClass
        {
            public static readonly Declaration Global = new Declaration(new QualifiedMemberName(VbaModuleName, "Global"), VbaLib.Vba, "VBA", "Global", false, false, Accessibility.Global, DeclarationType.Class);
            public static Declaration Load = new Declaration(new QualifiedMemberName(VbaModuleName, "Load"), Global, "VBA.Global", null, false, false, Accessibility.Public, DeclarationType.Procedure);
            public static Declaration Unload = new Declaration(new QualifiedMemberName(VbaModuleName, "Unload"), Global, "VBA.Global", null, false, false, Accessibility.Public, DeclarationType.Procedure);

            public static Declaration UserForms = new Declaration(new QualifiedMemberName(VbaModuleName, "UserForms"), Global, "VBA.Global", "Object", true, false, Accessibility.Public, DeclarationType.PropertyGet);
        }
        
        #endregion

        #region MSForms library (just for form events)
        /*
         *  This part should be deleted and Rubberduck should use MsFormsLib instead.
         *  However MsFormsLib is daunting and not implemented yet, and all we want for now
         *  is a Declaration object for form events - so this is "good enough" until MsFormsLib is implemented.
         */
        private static readonly QualifiedModuleName MsFormsModuleName = new QualifiedModuleName("MSForms", "MSForms");
        public static Declaration MsForms = new Declaration(new QualifiedMemberName(MsFormsModuleName, "MSForms"), null, "MSForms", "MSForms", false, false, Accessibility.Global, DeclarationType.Module);

        private class UserFormClass
        {
            public static readonly Declaration UserForm = new Declaration(new QualifiedMemberName(MsFormsModuleName, "UserForm"), MsForms, "MSForms", "UserForm", true, false, Accessibility.Global, DeclarationType.Class);

            // events
            public static Declaration AddControl = new Declaration(new QualifiedMemberName(MsFormsModuleName, "AddControl"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration BeforeDragOver = new Declaration(new QualifiedMemberName(MsFormsModuleName, "BeforeDragOver"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration BeforeDropOrPaste = new Declaration(new QualifiedMemberName(MsFormsModuleName, "BeforeDropOrPaste"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Click = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Click"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration DblClick = new Declaration(new QualifiedMemberName(MsFormsModuleName, "DblClick"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Error = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Error"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration KeyDown = new Declaration(new QualifiedMemberName(MsFormsModuleName, "KeyDown"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration KeyPress = new Declaration(new QualifiedMemberName(MsFormsModuleName, "KeyPress"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration KeyUp = new Declaration(new QualifiedMemberName(MsFormsModuleName, "KeyUp"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Layout = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Layout"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration MouseDown = new Declaration(new QualifiedMemberName(MsFormsModuleName, "MouseDown"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration MouseMove = new Declaration(new QualifiedMemberName(MsFormsModuleName, "MouseMove"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration MouseUp = new Declaration(new QualifiedMemberName(MsFormsModuleName, "MouseUp"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration RemoveControl = new Declaration(new QualifiedMemberName(MsFormsModuleName, "RemoveControl"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Scroll = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Scroll"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Zoom = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Zoom"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);

            // ghost events (nowhere in the object browser)
            public static Declaration Activate = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Activate"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Deactivate = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Deactivate"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Initialize = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Initialize"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration QueryClose = new Declaration(new QualifiedMemberName(MsFormsModuleName, "QueryClose"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Resize = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Resize"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Terminate = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Terminate"), UserForm, "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
        }
        #endregion
    }
}