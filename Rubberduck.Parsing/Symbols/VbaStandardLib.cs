using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces;
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
        private static readonly IRubberduckFactory<IRubberduckCodePane> Factory = new RubberduckCodePaneFactory();

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
            public static Declaration Vba = new Declaration(new QualifiedMemberName(VbaModuleName, "VBA"), "VBA", "VBA", true, false, Accessibility.Global, DeclarationType.Project, Factory);

            public static Declaration FormShowConstants = new Declaration(new QualifiedMemberName(VbaModuleName, "FormShowConstants"), "VBA", "FormShowConstants", false, false, Accessibility.Global, DeclarationType.Enumeration, Factory);
            public static Declaration VbModal = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbModal"), "VBA", "FormShowConstants", Accessibility.Global, DeclarationType.EnumerationMember, "1", Factory);
            public static Declaration VbModeless = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbModeless"), "VBA", "FormShowConstants", Accessibility.Global, DeclarationType.EnumerationMember, "0", Factory);

            public static Declaration VbAppWinStyle = new Declaration(new QualifiedMemberName(VbaModuleName, "VbAppWinStyle"), "VBA", "VbAppWinStyle", false, false, Accessibility.Global, DeclarationType.Enumeration, Factory);
            public static Declaration VbHide = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbHide"), "VBA", "VbAppWinStyle", Accessibility.Global, DeclarationType.EnumerationMember, "0", Factory);
            public static Declaration VbMaximizedFocus = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMaximizedFocus"), "VBA", "VbAppWinStyle", Accessibility.Global, DeclarationType.EnumerationMember, "3", Factory);
            public static Declaration VbMinimizedFocus = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMinimizedFocus"), "VBA", "VbAppWinStyle", Accessibility.Global, DeclarationType.EnumerationMember, "2", Factory);
            public static Declaration VbMinimizedNoFocus = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMinimizedNoFocus"), "VBA", "VbAppWinStyle", Accessibility.Global, DeclarationType.EnumerationMember, "6", Factory);
            public static Declaration VbNormalFocus = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbNormalFocus"), "VBA", "VbAppWinStyle", Accessibility.Global, DeclarationType.EnumerationMember, "1", Factory);
            public static Declaration VbNormalNoFocus = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbNormalNoFocus"), "VBA", "VbAppWinStyle", Accessibility.Global, DeclarationType.EnumerationMember, "4", Factory);

            public static Declaration VbCalendar = new Declaration(new QualifiedMemberName(VbaModuleName, "VbCalendar"), "VBA", "VbCalendar", false, false, Accessibility.Global, DeclarationType.Enumeration, Factory);
            public static Declaration VbCalGreg = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbCalGreg"), "VBA", "VbCalendar", Accessibility.Global, DeclarationType.EnumerationMember, "0", Factory);
            public static Declaration VbCalHijri = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbCalHijri"), "VBA", "VbCalendar", Accessibility.Global, DeclarationType.EnumerationMember, "1", Factory);

            public static Declaration VbCallType = new Declaration(new QualifiedMemberName(VbaModuleName, "VbCallType"), "VBA", "VbCallType", false, false, Accessibility.Global, DeclarationType.Enumeration, Factory);
            public static Declaration VbGet = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbGet"), "VBA", "VbCallType", Accessibility.Global, DeclarationType.EnumerationMember, "2", Factory);
            public static Declaration VbLet = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbLet"), "VBA", "VbCallType", Accessibility.Global, DeclarationType.EnumerationMember, "4", Factory);
            public static Declaration VbMethod = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMethod"), "VBA", "VbCallType", Accessibility.Global, DeclarationType.EnumerationMember, "1", Factory);
            public static Declaration VbSet = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbSet"), "VBA", "VbCallType", Accessibility.Global, DeclarationType.EnumerationMember, "8", Factory);

            public static Declaration VbCompareMethod = new Declaration(new QualifiedMemberName(VbaModuleName, "VbCompareMethod"), "VBA", "VbCompareMethod", false, false, Accessibility.Global, DeclarationType.Enumeration, Factory);
            public static Declaration VbBinaryCompare = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbBinaryCompare"), "VBA", "VbCompareMethod", Accessibility.Global, DeclarationType.EnumerationMember, "0", Factory);
            public static Declaration VbTextCompare = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbTextCompare"), "VBA", "VbCompareMethod", Accessibility.Global, DeclarationType.EnumerationMember, "1", Factory);

            public static Declaration VbDateTimeFormat = new Declaration(new QualifiedMemberName(VbaModuleName, "VbDateTimeFormat"), "VBA", "VbDateTimeFormat", false, false, Accessibility.Global, DeclarationType.Enumeration, Factory);
            public static Declaration VbGeneralDate = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbGeneralDate"), "VBA", "VbDateTimeFormat", Accessibility.Global, DeclarationType.EnumerationMember, "0", Factory);
            public static Declaration VbLongDate = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbLongDate"), "VBA", "VbDateTimeFormat", Accessibility.Global, DeclarationType.EnumerationMember, "1", Factory);
            public static Declaration VbLongTime = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbLongTime"), "VBA", "VbDateTimeFormat", Accessibility.Global, DeclarationType.EnumerationMember, "3", Factory);
            public static Declaration VbShortDate = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbShortDate"), "VBA", "VbDateTimeFormat", Accessibility.Global, DeclarationType.EnumerationMember, "2", Factory);
            public static Declaration VbShortTime = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbShortTime"), "VBA", "VbDateTimeFormat", Accessibility.Global, DeclarationType.EnumerationMember, "4", Factory);

            public static Declaration VbDayOfWeek = new Declaration(new QualifiedMemberName(VbaModuleName, "VbDayOfWeek"), "VBA", "VbDayOfWeek", false, false, Accessibility.Global, DeclarationType.Enumeration, Factory);
            public static Declaration VbFriday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFriday"), "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "6", Factory);
            public static Declaration VbMonday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMonday"), "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "2", Factory);
            public static Declaration VbSaturday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbSaturday"), "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "7", Factory);
            public static Declaration VbSunday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbSunday"), "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "1", Factory);
            public static Declaration VbThursday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbThursday"), "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "5", Factory);
            public static Declaration VbTuesday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbTuesday"), "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "3", Factory);
            public static Declaration VbUseSystemDayOfWeek = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbUseSystemDayOfWeek"), "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "0", Factory);
            public static Declaration VbWednesday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbWednesday"), "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "4", Factory);

            public static Declaration VbFileAttribute = new Declaration(new QualifiedMemberName(VbaModuleName, "VbFileAttribute"), "VBA", "VbFileAttribute", false, false, Accessibility.Global, DeclarationType.Enumeration, Factory);
            public static Declaration VbNormal = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbNormal"), "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "0", Factory);
            public static Declaration VbReadOnly = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbReadOnly"), "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "1", Factory);
            public static Declaration VbHidden = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbHidden"), "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "2", Factory);
            public static Declaration VbSystem = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbSystem"), "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "4", Factory);
            public static Declaration VbVolume = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbVolume"), "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "8", Factory);
            public static Declaration VbDirectory = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDirectory"), "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "16", Factory);
            public static Declaration VbArchive = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbArchive"), "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "32", Factory);
            public static Declaration VbAlias = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbAlias"), "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "64", Factory);

            public static Declaration VbFirstWeekOfYear = new Declaration(new QualifiedMemberName(VbaModuleName, "VbFirstWeekOfYear"), "VBA", "VbFirstWeekOfYear", false, false, Accessibility.Global, DeclarationType.Enumeration, Factory);
            public static Declaration VbFirstFourDays = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFirstFourDays"), "VBA", "VbFirstWeekOfYear", Accessibility.Global, DeclarationType.EnumerationMember, "2", Factory);
            public static Declaration VbFirstFullWeek = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFirstFullWeek"), "VBA", "VbFirstWeekOfYear", Accessibility.Global, DeclarationType.EnumerationMember, "3", Factory);
            public static Declaration VbFirstJan1 = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFirstJan1"), "VBA", "VbFirstWeekOfYear", Accessibility.Global, DeclarationType.EnumerationMember, "1", Factory);
            public static Declaration VbUseSystem = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbUseSystem"), "VBA", "VbFirstWeekOfYear", Accessibility.Global, DeclarationType.EnumerationMember, "0", Factory);

            public static Declaration VbIMEStatus = new Declaration(new QualifiedMemberName(VbaModuleName, "VbIMEStatus"), "VBA", "VbIMEStatus", false, false, Accessibility.Global, DeclarationType.Enumeration, Factory);
            public static Declaration VbIMEAlphaDbl = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEAlphaDbl"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "7", Factory);
            public static Declaration VbIMEAlphaSng = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEAlphaSng"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "8", Factory);
            public static Declaration VbIMEDisable = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEDisable"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "3", Factory);
            public static Declaration VbIMEHiragana = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEHiragana"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "4", Factory);
            public static Declaration VbIMEKatakanaDbl = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEKatakanaDbl"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "5", Factory);
            public static Declaration VbIMEKatakanaSng = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEKatakanaSng"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "6", Factory);
            public static Declaration VbIMEModeAlpha = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeAlpha"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "8", Factory);
            public static Declaration VbIMEModeAlphaFull = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeAlphaFull"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "7", Factory);
            public static Declaration VbIMEModeDisable = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeDisable"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "3", Factory);
            public static Declaration VbIMEModeHangul = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeHangul"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "10", Factory);
            public static Declaration VbIMEModeHangulFull = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeHangulFull"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "9", Factory);
            public static Declaration VbIMEModeHiragana = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeHiragana"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "4", Factory);
            public static Declaration VbIMEModeKatakana = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeKatakana"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "5", Factory);
            public static Declaration VbIMEModeKatakanaHalf = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeKatakanaHalf"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "6", Factory);
            public static Declaration VbIMEModeNoControl = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeNoControl"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember,"0", Factory);
            public static Declaration VbIMEModeOff = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeOff"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "2", Factory);
            public static Declaration VbIMEModeOn = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeOn"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "1", Factory);
            public static Declaration VbIMENoOp = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMENoOp"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "0", Factory);
            public static Declaration VbIMEOff = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEOff"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "2", Factory);
            public static Declaration VbIMEOn = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEOn"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "1", Factory);

            public static Declaration VbMsgBoxResult = new Declaration(new QualifiedMemberName(VbaModuleName, "VbMsgBoxResult"), "VBA", "VbMsgBoxResult", false, false, Accessibility.Global, DeclarationType.Enumeration, Factory);
            public static Declaration VbAbort = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbAbort"), "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "3", Factory);
            public static Declaration VbCancel = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbCancel"), "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "2", Factory);
            public static Declaration VbIgnore = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIgnore"), "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "5", Factory);
            public static Declaration VbNo = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbNo"), "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "7", Factory);
            public static Declaration VbOk = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbOk"), "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "1", Factory);
            public static Declaration VbRetry = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbRetry"), "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "4", Factory);
            public static Declaration VbYes = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbYes"), "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "6", Factory);

            public static Declaration VbMsgBoxStyle = new Declaration(new QualifiedMemberName(VbaModuleName, "VbMsgBoxStyle"), "VBA", "VbMsgBoxStyle", false, false, Accessibility.Global, DeclarationType.Enumeration, Factory);
            public static Declaration VbAbortRetryIgnore = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbAbortRetryIgnore"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "2", Factory);
            public static Declaration VbApplicationModal = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbApplicationModal"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "0", Factory);
            public static Declaration VbCritical = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbCritical"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "16", Factory);
            public static Declaration VbDefaultButton1 = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDefaultButton1"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "0", Factory);
            public static Declaration VbDefaultButton2 = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDefaultButton2"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "256", Factory);
            public static Declaration VbDefaultButton3 = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDefaultButton3"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "512", Factory);
            public static Declaration VbDefaultButton4 = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDefaultButton4"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "768", Factory);
            public static Declaration VbExclamation = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbExclamation"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "48", Factory);
            public static Declaration VbInformation = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbInformation"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "64", Factory);
            public static Declaration VbMsgBoxHelpButton = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMsgBoxHelpButton"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "16384", Factory);
            public static Declaration VbMsgBoxRight = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMsgBoxRight"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "524288", Factory);
            public static Declaration VbMsgBoxRtlReading = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMsgBoxRtlReading"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "1048576", Factory);
            public static Declaration VbMsgBoxSetForeground = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMsgBoxSetForeground"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "65536", Factory);
            public static Declaration VbOkCancel = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbOkCancel"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "1", Factory);
            public static Declaration VbOkOnly = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbOkOnly"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "0", Factory);
            public static Declaration VbQuestion = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbQuestion"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "32", Factory);
            public static Declaration VbRetryCancel = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbRetryCancel"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "5", Factory);
            public static Declaration VbSystemModal = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbSystemModal"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "4096", Factory);
            public static Declaration VbYesNo = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbYesNo"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "4", Factory);
            public static Declaration VbYesNoCancel = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbYesNoCancel"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "3", Factory);

            public static Declaration VbQueryClose = new Declaration(new QualifiedMemberName(VbaModuleName, "VbQueryClose"), "VBA", "VbQueryClose", false, false, Accessibility.Global, DeclarationType.Enumeration, Factory);
            public static Declaration VbAppTaskManager = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbAppTaskManager"), "VBA", "VbQueryClose", Accessibility.Global, DeclarationType.EnumerationMember, "3", Factory);
            public static Declaration VbAppWindows = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbAppWindows"), "VBA", "VbQueryClose", Accessibility.Global, DeclarationType.EnumerationMember, "2", Factory);
            public static Declaration VbFormCode = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFormCode"), "VBA", "VbQueryClose", Accessibility.Global, DeclarationType.EnumerationMember, "1", Factory);
            public static Declaration VbFormControlMenu = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFormControlMenu"), "VBA", "VbQueryClose", Accessibility.Global, DeclarationType.EnumerationMember, "0", Factory);
            public static Declaration VbFormMDIForm = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFormMDIForm"), "VBA", "VbQueryClose", Accessibility.Global, DeclarationType.EnumerationMember, "4", Factory);

            public static Declaration VbStrConv = new Declaration(new QualifiedMemberName(VbaModuleName, "VbStrConv"), "VBA", "VbStrConv", false, false, Accessibility.Global, DeclarationType.Enumeration, Factory);
            public static Declaration VbFromUnicode = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFromUnicode"), "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "128", Factory);
            public static Declaration VbHiragana = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbHiragana"), "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "32", Factory);
            public static Declaration VbKatakana = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbKatakana"), "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "16", Factory);
            public static Declaration VbLowerCase = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbLowerCase"), "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "2", Factory);
            public static Declaration VbNarrow = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbNarrow"), "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "8", Factory);
            public static Declaration VbProperCase = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbProperCase"), "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "3", Factory);
            public static Declaration VbUnicode = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbUnicode"), "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "64", Factory);
            public static Declaration VbUpperCase = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbUpperCase"), "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "1", Factory);
            public static Declaration VbWide = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbWide"), "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "4", Factory);

            public static Declaration VbTriState = new Declaration(new QualifiedMemberName(VbaModuleName, "VbTriState"), "VBA", "VbTriState", false, false, Accessibility.Global, DeclarationType.Enumeration, Factory);
            public static Declaration VbFalse = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFalse"), "VBA", "VbTriState", Accessibility.Global, DeclarationType.EnumerationMember, "0", Factory);
            public static Declaration VbTrue = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbTrue"), "VBA", "VbTriState", Accessibility.Global, DeclarationType.EnumerationMember, "-1", Factory);
            public static Declaration VbUseDefault = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbUseDefault"), "VBA", "VbTriState", Accessibility.Global, DeclarationType.EnumerationMember, "-2", Factory);

            public static Declaration VbVarType = new Declaration(new QualifiedMemberName(VbaModuleName, "VbVarType"), "VBA", "VbVarType", false, false, Accessibility.Global, DeclarationType.Enumeration, Factory);
            public static Declaration VbArray = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbArray"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "8192", Factory);
            public static Declaration VbBoolean = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbBoolean"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "11", Factory);
            public static Declaration VbByte = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbByte"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "17", Factory);
            public static Declaration VbCurrency = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbCurrency"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "6", Factory);
            public static Declaration VbDataObject = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDataObject"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "13", Factory);
            public static Declaration VbDate = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDate"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "7", Factory);
            public static Declaration VbDecimal = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDecimal"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "14", Factory);
            public static Declaration VbDouble = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDouble"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "5", Factory);
            public static Declaration VbEmpty = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbEmpty"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "0", Factory);
            public static Declaration VbError = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbError"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "10", Factory);
            public static Declaration VbInteger = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbInteger"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "2", Factory);
            public static Declaration VbLong = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbLong"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "3", Factory);
            public static Declaration VbLongLong = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbLongLong"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "20", Factory);
            public static Declaration VbNull = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbNull"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "1", Factory);
            public static Declaration VbObject = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbObject"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "9", Factory);
            public static Declaration VbSingle = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbSingle"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "4", Factory);
            public static Declaration VbString = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbString"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "8", Factory);
            public static Declaration VbUserDefinedType = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbUserDefinedType"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "36", Factory);
            public static Declaration VbVariant = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbVariant"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "12", Factory);
        }

        #region Predefined standard/procedural modules

        private class ColorConstantsModule
        {
            private static QualifiedModuleName ColorConstantsModuleName = new QualifiedModuleName("VBA", "ColorConstants");
            public static Declaration ColorConstants = new Declaration(new QualifiedMemberName(ColorConstantsModuleName, "ColorConstants"), "VBA", "ColorConstants", false, false, Accessibility.Global, DeclarationType.Module, Factory);
            public static Declaration VbBlack = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbBlack"), "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "0", Factory);
            public static Declaration VbBlue = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbBlue"), "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "16711680", Factory);
            public static Declaration VbCyan = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbCyan"), "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "16776960", Factory);
            public static Declaration VbGreen = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbGreen"), "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "65280", Factory);
            public static Declaration VbMagenta = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbMagenta"), "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "16711935", Factory);
            public static Declaration VbRed = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbRed"), "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "255", Factory);
            public static Declaration VbWhite = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbWhite"), "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "16777215", Factory);
            public static Declaration VbYellow = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbYellow"), "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "65535", Factory);
        }

        private class ConstantsModule
        {
            private static QualifiedModuleName ConstantsModuleName = new QualifiedModuleName("VBA", "Constants");
            public static Declaration Constants = new Declaration(new QualifiedMemberName(ConstantsModuleName, "Constants"), "VBA", "Constants", false, false, Accessibility.Global, DeclarationType.Module, Factory);
            public static Declaration VbBack = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbBack"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbCr = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbCr"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbCrLf = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbCrLf"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbFormFeed = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbFormFeed"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbLf = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbLf"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbNewLine = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbNewLine"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbNullChar = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbNullChar"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbTab = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbTab"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbVerticalTab = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbVerticalTab"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbNullString = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbNullString"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbObjectError = new ValuedDeclaration(new QualifiedMemberName(ConstantsModuleName, "vbObjectError"), "VBA.Constants", "Long", Accessibility.Global, DeclarationType.Constant, "-2147221504", Factory);
        }

        private class ConversionModule
        {
            private static QualifiedModuleName ConversionModuleName = new QualifiedModuleName("VBA", "Conversion");
            public static Declaration Conversion = new Declaration(new QualifiedMemberName(ConversionModuleName, "Conversion"), "VBA", "Conversion", false, false, Accessibility.Global, DeclarationType.Module, Factory);
            public static Declaration CBool = new Declaration(new QualifiedMemberName(ConversionModuleName, "CBool"), "VBA.Conversion", "Boolean", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration CByte = new Declaration(new QualifiedMemberName(ConversionModuleName, "CByte"), "VBA.Conversion", "Byte", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration CCur = new Declaration(new QualifiedMemberName(ConversionModuleName, "CCur"), "VBA.Conversion", "Currency", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration CDate = new Declaration(new QualifiedMemberName(ConversionModuleName, "CDate"), "VBA.Conversion", "Date", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration CVDate = new Declaration(new QualifiedMemberName(ConversionModuleName, "CVDate"), "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration CDbl = new Declaration(new QualifiedMemberName(ConversionModuleName, "CDbl"), "VBA.Conversion", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration CDec = new Declaration(new QualifiedMemberName(ConversionModuleName, "CDec"), "VBA.Conversion", "Decimal", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration CInt = new Declaration(new QualifiedMemberName(ConversionModuleName, "CInt"), "VBA.Conversion", "Integer", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration CLng = new Declaration(new QualifiedMemberName(ConversionModuleName, "CLng"), "VBA.Conversion", "Long", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration CLngLng = new Declaration(new QualifiedMemberName(ConversionModuleName, "CLngLng"), "VBA.Conversion", "LongLong", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration CLngPtr = new Declaration(new QualifiedMemberName(ConversionModuleName, "CLngPtr"), "VBA.Conversion", "LongPtr", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration CSng = new Declaration(new QualifiedMemberName(ConversionModuleName, "CSng"), "VBA.Conversion", "Single", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration CStr = new Declaration(new QualifiedMemberName(ConversionModuleName, "CStr"), "VBA.Conversion", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration CVar = new Declaration(new QualifiedMemberName(ConversionModuleName, "CVar"), "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration CVErr = new Declaration(new QualifiedMemberName(ConversionModuleName, "CVErr"), "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Error = new Declaration(new QualifiedMemberName(ConversionModuleName, "Error"), "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration ErrorStr = new Declaration(new QualifiedMemberName(ConversionModuleName, "Error$"), "VBA.Conversion", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Fix = new Declaration(new QualifiedMemberName(ConversionModuleName, "Fix"), "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Hex = new Declaration(new QualifiedMemberName(ConversionModuleName, "Hex"), "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration HexStr = new Declaration(new QualifiedMemberName(ConversionModuleName, "Hex$"), "VBA.Conversion", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Int = new Declaration(new QualifiedMemberName(ConversionModuleName, "Int"), "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration MacID = new Declaration(new QualifiedMemberName(ConversionModuleName, "MacID"), "VBA.Conversion", "Long", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Oct = new Declaration(new QualifiedMemberName(ConversionModuleName, "Oct"), "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration OctStr = new Declaration(new QualifiedMemberName(ConversionModuleName, "Oct$"), "VBA.Conversion", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Str = new Declaration(new QualifiedMemberName(ConversionModuleName, "Str"), "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration StrStr = new Declaration(new QualifiedMemberName(ConversionModuleName, "Str$"), "VBA.Conversion", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Val = new Declaration(new QualifiedMemberName(ConversionModuleName, "Val"), "VBA.Conversion", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
        }

        private class DateTimeModule
        {
            private static QualifiedModuleName DateTimeModuleName = new QualifiedModuleName("VBA", "DateTime");
            // functions
            public static Declaration DateTime = new Declaration(new QualifiedMemberName(DateTimeModuleName, "DateTime"), "VBA", "DateTime", false, false, Accessibility.Global, DeclarationType.Module, Factory);
            public static Declaration DateAdd = new Declaration(new QualifiedMemberName(DateTimeModuleName, "DateAdd"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration DateDiff = new Declaration(new QualifiedMemberName(DateTimeModuleName, "DateDiff"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration DatePart = new Declaration(new QualifiedMemberName(DateTimeModuleName, "DatePart"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration DateSerial = new Declaration(new QualifiedMemberName(DateTimeModuleName, "DateSerial"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration DateValue = new Declaration(new QualifiedMemberName(DateTimeModuleName, "DateValue"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Day = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Day"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Hour = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Hour"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Minute = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Minute"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Month = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Month"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Second = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Second"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration TimeSerial = new Declaration(new QualifiedMemberName(DateTimeModuleName, "TimeSerial"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration TimeValue = new Declaration(new QualifiedMemberName(DateTimeModuleName, "TimeValue"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration WeekDay = new Declaration(new QualifiedMemberName(DateTimeModuleName, "WeekDay"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Year = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Year"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            // properties
            public static Declaration Calendar = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Calendar"), "VBA.DateTime", "vbCalendar", false, false, Accessibility.Global, DeclarationType.PropertyGet, Factory);
            public static Declaration Date = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Date"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.PropertyGet, Factory);
            public static Declaration DateStr = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Date$"), "VBA.DateTime", "String", false, false, Accessibility.Global, DeclarationType.PropertyGet, Factory);
            public static Declaration Now = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Now"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.PropertyGet, Factory);
            public static Declaration Time = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Time"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.PropertyGet, Factory);
            public static Declaration TimeStr = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Time$"), "VBA.DateTime", "String", false, false, Accessibility.Global, DeclarationType.PropertyGet, Factory);
            public static Declaration Timer = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Timer"), "VBA.DateTime", "Single", false, false, Accessibility.Global, DeclarationType.PropertyGet, Factory);
        }

        private class FileSystemModule
        {
            private static QualifiedModuleName FileSystemModuleName = new QualifiedModuleName("VBA", "FileSystem");
            // functions
            public static Declaration FileSystem = new Declaration(new QualifiedMemberName(FileSystemModuleName, "FileSystem"), "VBA", "FileSystem", false, false, Accessibility.Global, DeclarationType.Module, Factory);
            public static Declaration CurDir = new Declaration(new QualifiedMemberName(FileSystemModuleName, "CurDir"), "VBA.FileSystem", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration CurDirStr = new Declaration(new QualifiedMemberName(FileSystemModuleName, "CurDir$"), "VBA.FileSystem", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Dir = new Declaration(new QualifiedMemberName(FileSystemModuleName, "Dir"), "VBA.FileSystem", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration EOF = new Declaration(new QualifiedMemberName(FileSystemModuleName, "EOF"), "VBA.FileSystem", "Boolean", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration FileAttr = new Declaration(new QualifiedMemberName(FileSystemModuleName, "FileAttr"), "VBA.FileSystem", "Long", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration FileDateTime = new Declaration(new QualifiedMemberName(FileSystemModuleName, "FileDateTime"), "VBA.FileSystem", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration FileLen = new Declaration(new QualifiedMemberName(FileSystemModuleName, "FileLen"), "VBA.FileSystem", "Long", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration FreeFile = new Declaration(new QualifiedMemberName(FileSystemModuleName, "FreeFile"), "VBA.FileSystem", "Integer", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Loc = new Declaration(new QualifiedMemberName(FileSystemModuleName, "Loc"), "VBA.FileSystem", "Long", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration LOF = new Declaration(new QualifiedMemberName(FileSystemModuleName, "LOF"), "VBA.FileSystem", "Long", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Seek = new Declaration(new QualifiedMemberName(FileSystemModuleName, "Seek"), "VBA.FileSystem", "Long", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            // procedures
            public static Declaration ChDir = new Declaration(new QualifiedMemberName(FileSystemModuleName, "ChDir"), "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure, Factory);
            public static Declaration ChDrive = new Declaration(new QualifiedMemberName(FileSystemModuleName, "ChDrive"), "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure, Factory);
            public static Declaration FileCopy = new Declaration(new QualifiedMemberName(FileSystemModuleName, "FileCopy"), "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure, Factory);
            public static Declaration Kill = new Declaration(new QualifiedMemberName(FileSystemModuleName, "Kill"), "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure, Factory);
            public static Declaration MkDir = new Declaration(new QualifiedMemberName(FileSystemModuleName, "MkDir"), "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure, Factory);
            public static Declaration RmDir = new Declaration(new QualifiedMemberName(FileSystemModuleName, "RmDir"), "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure, Factory);
            public static Declaration SetAttr = new Declaration(new QualifiedMemberName(FileSystemModuleName, "SetAttr"), "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure, Factory);
        }

        private class FinancialModule
        {
            private static QualifiedModuleName FinancialModuleName = new QualifiedModuleName("VBA", "Financial");
            public static Declaration Financial = new Declaration(new QualifiedMemberName(FinancialModuleName, "Financial"), "VBA", "Financial", false, false, Accessibility.Global, DeclarationType.Module, Factory);
            public static Declaration DDB = new Declaration(new QualifiedMemberName(FinancialModuleName, "DDB"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration FV = new Declaration(new QualifiedMemberName(FinancialModuleName, "FV"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration IPmt = new Declaration(new QualifiedMemberName(FinancialModuleName, "IPmt"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration IRR = new Declaration(new QualifiedMemberName(FinancialModuleName, "IRR"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration MIRR = new Declaration(new QualifiedMemberName(FinancialModuleName, "MIRR"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration NPer = new Declaration(new QualifiedMemberName(FinancialModuleName, "NPer"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration NPV = new Declaration(new QualifiedMemberName(FinancialModuleName, "NPV"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Pmt = new Declaration(new QualifiedMemberName(FinancialModuleName, "Pmt"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration PPmt = new Declaration(new QualifiedMemberName(FinancialModuleName, "PPmt"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration PV = new Declaration(new QualifiedMemberName(FinancialModuleName, "PV"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Rate = new Declaration(new QualifiedMemberName(FinancialModuleName, "Rate"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration SLN = new Declaration(new QualifiedMemberName(FinancialModuleName, "SLN"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration SYD = new Declaration(new QualifiedMemberName(FinancialModuleName, "SYD"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
        }

        private class HiddenModule
        {
            private static QualifiedModuleName HiddenModuleName = new QualifiedModuleName("VBA", "[_HiddenModule]");
            public static Declaration Array = new Declaration(new QualifiedMemberName(HiddenModuleName, "Array"), "VBA.[_HiddenModule]", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Input = new Declaration(new QualifiedMemberName(HiddenModuleName, "Input"), "VBA.[_HiddenModule]", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration InputStr = new Declaration(new QualifiedMemberName(HiddenModuleName, "Input$"), "VBA.[_HiddenModule]", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration InputB = new Declaration(new QualifiedMemberName(HiddenModuleName, "InputB"), "VBA.[_HiddenModule]", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration InputBStr = new Declaration(new QualifiedMemberName(HiddenModuleName, "InputB$"), "VBA.[_HiddenModule]", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Width = new Declaration(new QualifiedMemberName(HiddenModuleName, "Width"), "VBA.[_HiddenModule]", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            
            // hidden members... of hidden module (like, very very hidden!)
            public static Declaration ObjPtr = new Declaration(new QualifiedMemberName(HiddenModuleName, "ObjPtr"), "VBA.[_HiddenModule]", "LongPtr", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration StrPtr = new Declaration(new QualifiedMemberName(HiddenModuleName, "StrPtr"), "VBA.[_HiddenModule]", "LongPtr", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration VarPtr = new Declaration(new QualifiedMemberName(HiddenModuleName, "VarPtr"), "VBA.[_HiddenModule]", "LongPtr", false, false, Accessibility.Global, DeclarationType.Function, Factory);
        }

        private class InformationModule
        {
            private static QualifiedModuleName InformationModuleName = new QualifiedModuleName("VBA", "Information");
            public static Declaration Information = new Declaration(new QualifiedMemberName(InformationModuleName, "Information"), "VBA", "Information", false, false, Accessibility.Global, DeclarationType.Module, Factory);
            public static Declaration Err = new Declaration(new QualifiedMemberName(InformationModuleName, "Err"), "VBA.Information", "ErrObject", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Erl = new Declaration(new QualifiedMemberName(InformationModuleName, "Erl"), "VBA.Information", "Long", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration IMEStatus = new Declaration(new QualifiedMemberName(InformationModuleName, "IMEStatus"), "VBA.Information", "vbIMEStatus", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration IsArray = new Declaration(new QualifiedMemberName(InformationModuleName, "IsArray"), "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration IsDate = new Declaration(new QualifiedMemberName(InformationModuleName, "IsDate"), "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration IsEmpty = new Declaration(new QualifiedMemberName(InformationModuleName, "IsEmpty"), "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration IsError = new Declaration(new QualifiedMemberName(InformationModuleName, "IsError"), "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration IsMissing = new Declaration(new QualifiedMemberName(InformationModuleName, "IsMissing"), "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration IsNull = new Declaration(new QualifiedMemberName(InformationModuleName, "IsNull"), "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration IsNumeric = new Declaration(new QualifiedMemberName(InformationModuleName, "IsNumeric"), "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration IsObject = new Declaration(new QualifiedMemberName(InformationModuleName, "IsObject"), "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration QBColor = new Declaration(new QualifiedMemberName(InformationModuleName, "QBColor"), "VBA.Information", "Long", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration RGB = new Declaration(new QualifiedMemberName(InformationModuleName, "RGB"), "VBA.Information", "Long", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration TypeName = new Declaration(new QualifiedMemberName(InformationModuleName, "TypeName"), "VBA.Information", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration VarType = new Declaration(new QualifiedMemberName(InformationModuleName, "VarType"), "VBA.Information", "vbVarType", false, false, Accessibility.Global, DeclarationType.Function, Factory);
        }

        private class InteractionModule
        {
            private static QualifiedModuleName InteractionModuleName = new QualifiedModuleName("VBA", "Interaction");
            // functions
            public static Declaration Interaction = new Declaration(new QualifiedMemberName(InteractionModuleName, "Interaction"), "VBA", "Interaction", false, false, Accessibility.Global, DeclarationType.Module, Factory);
            public static Declaration CallByName = new Declaration(new QualifiedMemberName(InteractionModuleName, "CallByName"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Choose = new Declaration(new QualifiedMemberName(InteractionModuleName, "Choose"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Command = new Declaration(new QualifiedMemberName(InteractionModuleName, "Command"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration CommandStr = new Declaration(new QualifiedMemberName(InteractionModuleName, "Command$"), "VBA.Interaction", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration CreateObject = new Declaration(new QualifiedMemberName(InteractionModuleName, "CreateObject"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration DoEvents = new Declaration(new QualifiedMemberName(InteractionModuleName, "DoEvents"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Environ = new Declaration(new QualifiedMemberName(InteractionModuleName, "Environ"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration EnvironStr = new Declaration(new QualifiedMemberName(InteractionModuleName, "Environ$"), "VBA.Interaction", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration GetAllSettings = new Declaration(new QualifiedMemberName(InteractionModuleName, "GetAllSettings"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration GetAttr = new Declaration(new QualifiedMemberName(InteractionModuleName, "GetAttr"), "VBA.Interaction", "vbFileAttribute", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration GetObject = new Declaration(new QualifiedMemberName(InteractionModuleName, "GetObject"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration GetSetting = new Declaration(new QualifiedMemberName(InteractionModuleName, "GetSetting"), "VBA.Interaction", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration IIf = new Declaration(new QualifiedMemberName(InteractionModuleName, "IIf"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration InputBox = new Declaration(new QualifiedMemberName(InteractionModuleName, "InputBox"), "VBA.Interaction", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration MacScript = new Declaration(new QualifiedMemberName(InteractionModuleName, "MacScript"), "VBA.Interaction", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration MsgBox = new Declaration(new QualifiedMemberName(InteractionModuleName, "MsgBox"), "VBA.Interaction", "vbMsgBoxResult", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Partition = new Declaration(new QualifiedMemberName(InteractionModuleName, "Partition"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Shell = new Declaration(new QualifiedMemberName(InteractionModuleName, "Shell"), "VBA.Interaction", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Switch = new Declaration(new QualifiedMemberName(InteractionModuleName, "Switch"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            // procedures
            public static Declaration AppActivate = new Declaration(new QualifiedMemberName(InteractionModuleName, "AppActivate"), "VBA.Interaction", null, false, false, Accessibility.Global, DeclarationType.Procedure, Factory);
            public static Declaration Beep = new Declaration(new QualifiedMemberName(InteractionModuleName, "Beep"), "VBA.Interaction", null, false, false, Accessibility.Global, DeclarationType.Procedure, Factory);
            public static Declaration DeleteSetting = new Declaration(new QualifiedMemberName(InteractionModuleName, "DeleteSetting"), "VBA.Interaction", null, false, false, Accessibility.Global, DeclarationType.Procedure, Factory);
            public static Declaration SaveSetting = new Declaration(new QualifiedMemberName(InteractionModuleName, "SaveSetting"), "VBA.Interaction", null, false, false, Accessibility.Global, DeclarationType.Procedure, Factory);
            public static Declaration SendKeys = new Declaration(new QualifiedMemberName(InteractionModuleName, "SendKeys"), "VBA.Interaction", null, false, false, Accessibility.Global, DeclarationType.Procedure, Factory);
        }

        private class KeyCodeConstantsModule
        {
            private static QualifiedModuleName KeyCodeConstantsModuleName = new QualifiedModuleName("VBA", "KeyCodeConstants");
            public static Declaration KeyCodeConstants = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "KeyCodeConstants"), "VBA", "KeyCodeConstants", false, false, Accessibility.Global, DeclarationType.Module, Factory);
            public static Declaration VbKeyLButton = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyLButton"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "1", Factory);
            public static Declaration VbKeyRButton = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyRButton"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "2", Factory);
            public static Declaration VbKeyCancel = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyCancel"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "3", Factory);
            public static Declaration VbKeyMButton = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyMButton"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "4", Factory);
            public static Declaration VbKeyBack = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyBack"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "8", Factory);
            public static Declaration VbKeyTab = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyTab"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "9", Factory);
            public static Declaration VbKeyClear = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyClear"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "12", Factory);
            public static Declaration VbKeyReturn = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyReturn"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "13", Factory);
            public static Declaration VbKeyShift = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyShift"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "16", Factory);
            public static Declaration VbKeyControl = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyControl"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "17", Factory);
            public static Declaration VbKeyMenu = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyMenu"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "18", Factory);
            public static Declaration VbKeyPause = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyPause"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "19", Factory);
            public static Declaration VbKeyCapital = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyCapital"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "20", Factory);
            public static Declaration VbKeyEscape = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyEscape"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "27", Factory);
            public static Declaration VbKeySpace = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeySpace"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "32", Factory);
            public static Declaration VbKeyPageUp = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyPageUp"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "33", Factory);
            public static Declaration VbKeyPageDown = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyPageDown"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "34", Factory);
            public static Declaration VbKeyEnd = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyEnd"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "35", Factory);
            public static Declaration VbKeyHome = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyHome"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "36", Factory);
            public static Declaration VbKeyLeft = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyLeft"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "37", Factory);
            public static Declaration VbKeyUp = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyUp"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "38", Factory);
            public static Declaration VbKeyRight = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyRight"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "39", Factory);
            public static Declaration VbKeyDown = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyDown"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "40", Factory);
            public static Declaration VbKeySelect = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeySelect"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "41", Factory);
            public static Declaration VbKeyPrint = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyPrint"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "42", Factory);
            public static Declaration VbKeyExecute = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyExecute"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "43", Factory);
            public static Declaration VbKeySnapshot = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeySnapshot"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "44", Factory);
            public static Declaration VbKeyInsert = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyInsert"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "45", Factory);
            public static Declaration VbKeyDelete = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyDelete"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "46", Factory);
            public static Declaration VbKeyHelp = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyHelp"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "47", Factory);
            public static Declaration VbKeyNumLock = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumLock"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "144", Factory);
            public static Declaration VbKeyA = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyA"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "65", Factory);
            public static Declaration VbKeyB = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyB"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "66", Factory);
            public static Declaration VbKeyC = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyC"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "67", Factory);
            public static Declaration VbKeyD = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyD"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "68", Factory);
            public static Declaration VbKeyE = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyE"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "69", Factory);
            public static Declaration VbKeyF = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "70", Factory);
            public static Declaration VbKeyG = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyG"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "71", Factory);
            public static Declaration VbKeyH = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyH"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "72", Factory);
            public static Declaration VbKeyI = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyI"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "73", Factory);
            public static Declaration VbKeyJ = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyJ"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "74", Factory);
            public static Declaration VbKeyK = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyK"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "75", Factory);
            public static Declaration VbKeyL = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyL"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "76", Factory);
            public static Declaration VbKeyM = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyM"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "77", Factory);
            public static Declaration VbKeyN = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyN"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "78", Factory);
            public static Declaration VbKeyO = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyO"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "79", Factory);
            public static Declaration VbKeyP = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyP"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "80", Factory);
            public static Declaration VbKeyQ = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyQ"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "81", Factory);
            public static Declaration VbKeyR = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyR"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "82", Factory);
            public static Declaration VbKeyS = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyS"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "83", Factory);
            public static Declaration VbKeyT = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyT"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "84", Factory);
            public static Declaration VbKeyU = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyU"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "85", Factory);
            public static Declaration VbKeyV = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyV"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "86", Factory);
            public static Declaration VbKeyW = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyW"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "87", Factory);
            public static Declaration VbKeyX = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyX"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "88", Factory);
            public static Declaration VbKeyY = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyY"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "89", Factory);
            public static Declaration VbKeyZ = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyZ"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "90", Factory);
            public static Declaration VbKey0 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey0"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "48", Factory);
            public static Declaration VbKey1 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey1"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "49", Factory);
            public static Declaration VbKey2 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey2"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "50", Factory);
            public static Declaration VbKey3 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey3"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "51", Factory);
            public static Declaration VbKey4 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey4"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "52", Factory);
            public static Declaration VbKey5 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey5"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "53", Factory);
            public static Declaration VbKey6 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey6"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "54", Factory);
            public static Declaration VbKey7 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey7"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "55", Factory);
            public static Declaration VbKey8 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey8"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "56", Factory);
            public static Declaration VbKey9 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey9"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "57", Factory);
            public static Declaration VbKeyNumpad0 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad0"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "96", Factory);
            public static Declaration VbKeyNumpad1 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad1"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "97", Factory);
            public static Declaration VbKeyNumpad2 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad2"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "98", Factory);
            public static Declaration VbKeyNumpad3 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad3"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "99", Factory);
            public static Declaration VbKeyNumpad4 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad4"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "100", Factory);
            public static Declaration VbKeyNumpad5 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad5"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "101", Factory);
            public static Declaration VbKeyNumpad6 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad6"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "102", Factory);
            public static Declaration VbKeyNumpad7 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad7"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "103", Factory);
            public static Declaration VbKeyNumpad8 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad8"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "104", Factory);
            public static Declaration VbKeyNumpad9 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad9"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "105", Factory);
            public static Declaration VbKeyMultiply = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyMultiply"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "106", Factory);
            public static Declaration VbKeyAdd = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyAdd"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "107", Factory);
            public static Declaration VbKeySeparator = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeySeparator"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "108", Factory);
            public static Declaration VbKeySubtract = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeySubtract"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "109", Factory);
            public static Declaration VbKeyDecimal = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyDecimal"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "110", Factory);
            public static Declaration VbKeyDivide = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyDivide"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "111", Factory);
            public static Declaration VbKeyF1 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF1"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "112", Factory);
            public static Declaration VbKeyF2 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF2"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "113", Factory);
            public static Declaration VbKeyF3 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF3"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "114", Factory);
            public static Declaration VbKeyF4 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF4"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "115", Factory);
            public static Declaration VbKeyF5 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF5"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "116", Factory);
            public static Declaration VbKeyF6 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF6"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "117", Factory);
            public static Declaration VbKeyF7 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF7"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "118", Factory);
            public static Declaration VbKeyF8 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF8"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "119", Factory);
            public static Declaration VbKeyF9 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF9"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "120", Factory);
            public static Declaration VbKeyF10 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF10"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "121", Factory);
            public static Declaration VbKeyF11 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF11"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "122", Factory);
            public static Declaration VbKeyF12 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF12"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "123", Factory);
            public static Declaration VbKeyF13 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF13"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "124", Factory);
            public static Declaration VbKeyF14 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF14"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "125", Factory);
            public static Declaration VbKeyF15 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF15"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "126", Factory);
            public static Declaration VbKeyF16 = new ValuedDeclaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF16"), "VBA.KeyCodeConstants", "Long", Accessibility.Global, DeclarationType.Constant, "127", Factory);
        }

        private class MathModule
        {
            private static QualifiedModuleName MathModuleName = new QualifiedModuleName("VBA", "Math");
            // functions
            public static Declaration Math = new Declaration(new QualifiedMemberName(MathModuleName, "Math"), "VBA", "Math", false, false, Accessibility.Global, DeclarationType.Module, Factory);
            public static Declaration Abs = new Declaration(new QualifiedMemberName(MathModuleName, "Abs"), "VBA.Math", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Atn = new Declaration(new QualifiedMemberName(MathModuleName, "Atn"), "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Cos = new Declaration(new QualifiedMemberName(MathModuleName, "Cos"), "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Exp = new Declaration(new QualifiedMemberName(MathModuleName, "Exp"), "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Log = new Declaration(new QualifiedMemberName(MathModuleName, "Log"), "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Rnd = new Declaration(new QualifiedMemberName(MathModuleName, "Rnd"), "VBA.Math", "Single", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Round = new Declaration(new QualifiedMemberName(MathModuleName, "Round"), "VBA.Math", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Sgn = new Declaration(new QualifiedMemberName(MathModuleName, "Sgn"), "VBA.Math", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Sin = new Declaration(new QualifiedMemberName(MathModuleName, "Sin"), "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Sqr = new Declaration(new QualifiedMemberName(MathModuleName, "Sqr"), "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Tan = new Declaration(new QualifiedMemberName(MathModuleName, "Tan"), "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            //procedures
            public static Declaration Randomize = new Declaration(new QualifiedMemberName(MathModuleName, "Randomize"), "VBA.Math", null, false, false, Accessibility.Global, DeclarationType.Procedure, Factory);
        }

        private class StringsModule
        {
            private static QualifiedModuleName StringsModuleName = new QualifiedModuleName("VBA", "Strings");
            public static Declaration Strings = new Declaration(new QualifiedMemberName(StringsModuleName, "Strings"), "VBA", "Strings", false, false, Accessibility.Global, DeclarationType.Module, Factory);
            public static Declaration Asc = new Declaration(new QualifiedMemberName(StringsModuleName, "Asc"), "VBA.Strings", "Integer", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration AscW = new Declaration(new QualifiedMemberName(StringsModuleName, "AscW"), "VBA.Strings", "Integer", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration AscB = new Declaration(new QualifiedMemberName(StringsModuleName, "AscB"), "VBA.Strings", "Integer", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Chr = new Declaration(new QualifiedMemberName(StringsModuleName, "Chr"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration ChrStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Chr$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration ChrB = new Declaration(new QualifiedMemberName(StringsModuleName, "ChrB"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration ChrBStr = new Declaration(new QualifiedMemberName(StringsModuleName, "ChrB$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration ChrW = new Declaration(new QualifiedMemberName(StringsModuleName, "ChrW"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration ChrWStr = new Declaration(new QualifiedMemberName(StringsModuleName, "ChrW$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Filter = new Declaration(new QualifiedMemberName(StringsModuleName, "Filter"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Format = new Declaration(new QualifiedMemberName(StringsModuleName, "Format"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration FormatStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Format$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration FormatCurrency = new Declaration(new QualifiedMemberName(StringsModuleName, "FormatCurrency"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration FormatDateTime = new Declaration(new QualifiedMemberName(StringsModuleName, "FormatDateTime"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration FormatNumber = new Declaration(new QualifiedMemberName(StringsModuleName, "FormatNumber"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration FormatPercent = new Declaration(new QualifiedMemberName(StringsModuleName, "FormatPercent"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration InStr = new Declaration(new QualifiedMemberName(StringsModuleName, "InStr"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration InStrB = new Declaration(new QualifiedMemberName(StringsModuleName, "InStrB"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration InStrRev = new Declaration(new QualifiedMemberName(StringsModuleName, "InStrRev"), "VBA.Strings", "Long", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Join = new Declaration(new QualifiedMemberName(StringsModuleName, "Join"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration LCase = new Declaration(new QualifiedMemberName(StringsModuleName, "LCase"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration LCaseStr = new Declaration(new QualifiedMemberName(StringsModuleName, "LCase$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Left = new Declaration(new QualifiedMemberName(StringsModuleName, "Left"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration LeftB = new Declaration(new QualifiedMemberName(StringsModuleName, "LeftB"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration LeftStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Left$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration LeftBStr = new Declaration(new QualifiedMemberName(StringsModuleName, "LeftB$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Len = new Declaration(new QualifiedMemberName(StringsModuleName, "Len"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration LenB = new Declaration(new QualifiedMemberName(StringsModuleName, "LenB"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration LTrim = new Declaration(new QualifiedMemberName(StringsModuleName, "LTrim"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration RTrim = new Declaration(new QualifiedMemberName(StringsModuleName, "RTrim"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Trim = new Declaration(new QualifiedMemberName(StringsModuleName, "Trim"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration LTrimStr = new Declaration(new QualifiedMemberName(StringsModuleName, "LTrim$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration RTrimStr = new Declaration(new QualifiedMemberName(StringsModuleName, "RTrim$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration TrimStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Trim$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Mid = new Declaration(new QualifiedMemberName(StringsModuleName, "Mid"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration MidB = new Declaration(new QualifiedMemberName(StringsModuleName, "MidB"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration MidStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Mid$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration MidBStr = new Declaration(new QualifiedMemberName(StringsModuleName, "MidB$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration MonthName = new Declaration(new QualifiedMemberName(StringsModuleName, "MonthName"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Replace = new Declaration(new QualifiedMemberName(StringsModuleName, "Replace"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Right = new Declaration(new QualifiedMemberName(StringsModuleName, "Right"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration RightB = new Declaration(new QualifiedMemberName(StringsModuleName, "RightB"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration RightStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Right$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration RightBStr = new Declaration(new QualifiedMemberName(StringsModuleName, "RightB$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Space = new Declaration(new QualifiedMemberName(StringsModuleName, "Space"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration SpaceStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Space$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration Split = new Declaration(new QualifiedMemberName(StringsModuleName, "Split"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration StrComp = new Declaration(new QualifiedMemberName(StringsModuleName, "StrComp"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration StrConv = new Declaration(new QualifiedMemberName(StringsModuleName, "StrConv"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration String = new Declaration(new QualifiedMemberName(StringsModuleName, "String"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration StringStr = new Declaration(new QualifiedMemberName(StringsModuleName, "String$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration StrReverse = new Declaration(new QualifiedMemberName(StringsModuleName, "StrReverse"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration UCase = new Declaration(new QualifiedMemberName(StringsModuleName, "UCase"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration UCaseStr = new Declaration(new QualifiedMemberName(StringsModuleName, "UCase$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
            public static Declaration WeekdayName = new Declaration(new QualifiedMemberName(StringsModuleName, "WeekdayName"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function, Factory);
        }

        private class SystemColorConstantsModule
        {
            private static QualifiedModuleName SystemColorConstantsModuleName = new QualifiedModuleName("VBA", "SystemColorConstants");
            public static Declaration SystemColorConstants = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "SystemColorConstants"), "VBA", "SystemColorConstants", false, false, Accessibility.Global, DeclarationType.Module, Factory);
            public static Declaration VbScrollBars = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbScrollBars"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbDesktop = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbDesktop"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbActiveTitleBar = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbActiveTitleBar"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbInactiveTitleBar = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbInactiveTitleBar"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbMenuBar = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbMenuBar"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbWindowBackground = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbWindowBackground"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbWindowFrame = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbWindowFrame"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbMenuText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbMenuText"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbWindowText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbWindowText"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbTitleBarText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbTitleBarText"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbActiveBorder = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbActiveBorder"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbInactiveBorder = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbInactiveBorder"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbApplicationWorkspace = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbApplicationWorkspace"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbHighlight = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbHighlight"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbHighlightText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbHighlightText"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbButtonFace = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbButtonFace"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbButtonShadow = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbButtonShadow"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbGrayText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbGrayText"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbButtonText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbButtonText"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbInactiveCaptionText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbInactiveCaptionText"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration Vb3DHighlight = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vb3DHighlight"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration Vb3DDKShadow = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vb3DDKShadow"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration Vb3DLight = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vb3DLight"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration Vb3DFace = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vb3DFace"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration Vb3DShadow = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vb3DShadow"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbInfoText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbInfoText"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
            public static Declaration VbInfoBackground = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbInfoBackground"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant, Factory);
        }

        #endregion

        #region Predefined class modules

        private class CollectionClass
        {
            public static Declaration Collection = new Declaration(new QualifiedMemberName(VbaModuleName, "Collection"), "VBA", "Collection", false, false, Accessibility.Global, DeclarationType.Class, Factory);
            public static Declaration NewEnum = new Declaration(new QualifiedMemberName(VbaModuleName, "[_NewEnum]"), "VBA.Collection", "Unknown", false, false, Accessibility.Public, DeclarationType.Function, Factory);
            public static Declaration Count = new Declaration(new QualifiedMemberName(VbaModuleName, "Count"), "VBA.Collection", "Long", false, false, Accessibility.Public, DeclarationType.Function, Factory);
            public static Declaration Item = new Declaration(new QualifiedMemberName(VbaModuleName, "Item"), "VBA.Collection", "Variant", false, false, Accessibility.Public, DeclarationType.Function, Factory);
            public static Declaration Add = new Declaration(new QualifiedMemberName(VbaModuleName, "Add"), "VBA.Collection", null, false, false, Accessibility.Public, DeclarationType.Procedure, Factory);
            public static Declaration Remove = new Declaration(new QualifiedMemberName(VbaModuleName, "Remove"), "VBA.Collection", null, false, false, Accessibility.Public, DeclarationType.Procedure, Factory);
        }

        private class ErrObjectClass
        {
            public static Declaration ErrObject = new Declaration(new QualifiedMemberName(VbaModuleName, "ErrObject"), "VBA", "ErrObject", false, false, Accessibility.Global, DeclarationType.Class, Factory);
            public static Declaration Clear = new Declaration(new QualifiedMemberName(VbaModuleName, "Clear"), "VBA.ErrObject", null, false, false, Accessibility.Public, DeclarationType.Procedure, Factory);
            public static Declaration Raise = new Declaration(new QualifiedMemberName(VbaModuleName, "Raise"), "VBA.ErrObject", null, false, false, Accessibility.Public, DeclarationType.Procedure, Factory);
            public static Declaration Description = new Declaration(new QualifiedMemberName(VbaModuleName, "Description"), "VBA.ErrObject", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet, Factory);
            public static Declaration HelpContext = new Declaration(new QualifiedMemberName(VbaModuleName, "HelpContext"), "VBA.ErrObject", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet, Factory);
            public static Declaration HelpFile = new Declaration(new QualifiedMemberName(VbaModuleName, "HelpFile"), "VBA.ErrObject", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet, Factory);
            public static Declaration LastDllError = new Declaration(new QualifiedMemberName(VbaModuleName, "LastDllError"), "VBA.ErrObject", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet, Factory);
            public static Declaration Number = new Declaration(new QualifiedMemberName(VbaModuleName, "Number"), "VBA.ErrObject", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet, Factory);
            public static Declaration Source = new Declaration(new QualifiedMemberName(VbaModuleName, "Source"), "VBA.ErrObject", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet, Factory);
        }

        private class GlobalClass
        {
            public static Declaration Global = new Declaration(new QualifiedMemberName(VbaModuleName, "Global"), "VBA", "Global", false, false, Accessibility.Global, DeclarationType.Class, Factory);
            public static Declaration Load = new Declaration(new QualifiedMemberName(VbaModuleName, "Load"), "VBA.Global", null, false, false, Accessibility.Public, DeclarationType.Procedure, Factory);
            public static Declaration Unload = new Declaration(new QualifiedMemberName(VbaModuleName, "Unload"), "VBA.Global", null, false, false, Accessibility.Public, DeclarationType.Procedure, Factory);

            public static Declaration UserForms = new Declaration(new QualifiedMemberName(VbaModuleName, "UserForms"), "VBA.Global", "Object", true, false, Accessibility.Public, DeclarationType.PropertyGet, Factory);
        }
        
        #endregion

        #region MSForms library (just for form events)
        /*
         *  This part should be deleted and Rubberduck should use MsFormsLib instead.
         *  However MsFormsLib is daunting and not implemented yet, and all we want for now
         *  is a Declaration object for form events - so this is "good enough" until MsFormsLib is implemented.
         */
        private static readonly QualifiedModuleName MsFormsModuleName = new QualifiedModuleName("MSForms", "MSForms");

        private class UserFormClass
        {
            public static Declaration UserForm = new Declaration(new QualifiedMemberName(MsFormsModuleName, "UserForm"), "MSForms", "UserForm", true, false, Accessibility.Global, DeclarationType.Class, Factory);

            // events
            public static Declaration AddControl = new Declaration(new QualifiedMemberName(MsFormsModuleName, "AddControl"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration BeforeDragOver = new Declaration(new QualifiedMemberName(MsFormsModuleName, "BeforeDragOver"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration BeforeDropOrPaste = new Declaration(new QualifiedMemberName(MsFormsModuleName, "BeforeDropOrPaste"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration Click = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Click"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration DblClick = new Declaration(new QualifiedMemberName(MsFormsModuleName, "DblClick"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration Error = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Error"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration KeyDown = new Declaration(new QualifiedMemberName(MsFormsModuleName, "KeyDown"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration KeyPress = new Declaration(new QualifiedMemberName(MsFormsModuleName, "KeyPress"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration KeyUp = new Declaration(new QualifiedMemberName(MsFormsModuleName, "KeyUp"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration Layout = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Layout"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration MouseDown = new Declaration(new QualifiedMemberName(MsFormsModuleName, "MouseDown"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration MouseMove = new Declaration(new QualifiedMemberName(MsFormsModuleName, "MouseMove"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration MouseUp = new Declaration(new QualifiedMemberName(MsFormsModuleName, "MouseUp"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration RemoveControl = new Declaration(new QualifiedMemberName(MsFormsModuleName, "RemoveControl"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration Scroll = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Scroll"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration Zoom = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Zoom"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);

            // ghost events (nowhere in the object browser)
            public static Declaration Activate = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Activate"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration Deactivate = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Deactivate"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration Initialize = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Initialize"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration QueryClose = new Declaration(new QualifiedMemberName(MsFormsModuleName, "QueryClose"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration Resize = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Resize"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
            public static Declaration Terminate = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Terminate"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event, Factory);
        }
        #endregion
    }
}