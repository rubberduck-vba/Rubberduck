using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Rubberduck.Parsing.Symbols
{
    /// <summary>
    /// Defines <see cref="Declaration"/> objects for the standard library.
    /// </summary>
    internal static class VbaStandardLib
    {
        private static IEnumerable<Declaration> StandardLibDeclarations;
        private static QualifiedModuleName VbaModuleName = new QualifiedModuleName("VBA", "VBA");

        public static IEnumerable<Declaration> Declarations
        {
            get
            {
                if (StandardLibDeclarations == null)
                {
                    var nestedTypes = typeof(VbaStandardLib).GetNestedTypes(BindingFlags.NonPublic);
                    var fields = nestedTypes.SelectMany(t => t.GetFields());
                    var values = fields.Select(f => f.GetValue(null));
                    StandardLibDeclarations = values.Cast<Declaration>();
                }

                return StandardLibDeclarations;
            }
        }

        private class VbaLib
        {
            public static Declaration Vba = new Declaration(new QualifiedMemberName(VbaModuleName, "VBA"), "VBA", "VBA", true, false, Accessibility.Global, DeclarationType.Project);

            public static Declaration FormShowConstants = new Declaration(new QualifiedMemberName(VbaModuleName, "FormShowConstants"), "VBA", "FormShowConstants", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbModal = new Declaration(new QualifiedMemberName(VbaModuleName, "vbModal"), "VBA", "FormShowConstants", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbModeless = new Declaration(new QualifiedMemberName(VbaModuleName, "vbModeless"), "VBA", "FormShowConstants", true, false, Accessibility.Global, DeclarationType.EnumerationMember);

            public static Declaration VbAppWinStyle = new Declaration(new QualifiedMemberName(VbaModuleName, "VbAppWinStyle"), "VBA", "VbAppWinStyle", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbHide = new Declaration(new QualifiedMemberName(VbaModuleName, "vbHide"), "VBA", "VbAppWinStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbMaximizedFocus = new Declaration(new QualifiedMemberName(VbaModuleName, "vbMaximizedFocus"), "VBA", "VbAppWinStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbMinimizedFocus = new Declaration(new QualifiedMemberName(VbaModuleName, "vbMinimizedFocus"), "VBA", "VbAppWinStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbMinimizedNoFocus = new Declaration(new QualifiedMemberName(VbaModuleName, "vbMinimizedNoFocus"), "VBA", "VbAppWinStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbNormalFocus = new Declaration(new QualifiedMemberName(VbaModuleName, "vbNormalFocus"), "VBA", "VbAppWinStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbNormalNoFocus = new Declaration(new QualifiedMemberName(VbaModuleName, "vbNormalNoFocus"), "VBA", "VbAppWinStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);

            public static Declaration VbCalendar = new Declaration(new QualifiedMemberName(VbaModuleName, "VbCalendar"), "VBA", "VbCalendar", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbCalGreg = new Declaration(new QualifiedMemberName(VbaModuleName, "vbCalGreg"), "VBA", "VbCalendar", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbCalHijri = new Declaration(new QualifiedMemberName(VbaModuleName, "vbCalHijri"), "VBA", "VbCalendar", true, false, Accessibility.Global, DeclarationType.EnumerationMember);

            public static Declaration VbCallType = new Declaration(new QualifiedMemberName(VbaModuleName, "VbCallType"), "VBA", "VbCallType", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbGet = new Declaration(new QualifiedMemberName(VbaModuleName, "vbGet"), "VBA", "VbCallType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbLet = new Declaration(new QualifiedMemberName(VbaModuleName, "vbLet"), "VBA", "VbCallType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbMethod = new Declaration(new QualifiedMemberName(VbaModuleName, "vbMethod"), "VBA", "VbCallType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbSet = new Declaration(new QualifiedMemberName(VbaModuleName, "vbSet"), "VBA", "VbCallType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);

            public static Declaration VbCompareMethod = new Declaration(new QualifiedMemberName(VbaModuleName, "VbCompareMethod"), "VBA", "VbCompareMethod", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbBinaryCompare = new Declaration(new QualifiedMemberName(VbaModuleName, "vbBinaryCompare"), "VBA", "VbCompareMethod", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbTextCompare = new Declaration(new QualifiedMemberName(VbaModuleName, "vbTextCompare"), "VBA", "VbCompareMethod", true, false, Accessibility.Global, DeclarationType.EnumerationMember);

            public static Declaration VbDateTimeFormat = new Declaration(new QualifiedMemberName(VbaModuleName, "VbDateTimeFormat"), "VBA", "VbDateTimeFormat", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbGeneralDate = new Declaration(new QualifiedMemberName(VbaModuleName, "vbGeneralDate"), "VBA", "VbDateTimeFormat", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbLongDate = new Declaration(new QualifiedMemberName(VbaModuleName, "vbLongDate"), "VBA", "VbDateTimeFormat", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbLongTime = new Declaration(new QualifiedMemberName(VbaModuleName, "vbLongTime"), "VBA", "VbDateTimeFormat", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbShortDate = new Declaration(new QualifiedMemberName(VbaModuleName, "vbShortDate"), "VBA", "VbDateTimeFormat", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbShortTime = new Declaration(new QualifiedMemberName(VbaModuleName, "vbShortTime"), "VBA", "VbDateTimeFormat", true, false, Accessibility.Global, DeclarationType.EnumerationMember);

            public static Declaration VbDayOfWeek = new Declaration(new QualifiedMemberName(VbaModuleName, "VbDayOfWeek"), "VBA", "VbDayOfWeek", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbFriday = new Declaration(new QualifiedMemberName(VbaModuleName, "vbFriday"), "VBA", "VbDayOfWeek", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbMonday = new Declaration(new QualifiedMemberName(VbaModuleName, "vbMonday"), "VBA", "VbDayOfWeek", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbSaturday = new Declaration(new QualifiedMemberName(VbaModuleName, "vbSaturday"), "VBA", "VbDayOfWeek", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbSunday = new Declaration(new QualifiedMemberName(VbaModuleName, "vbSunday"), "VBA", "VbDayOfWeek", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbThursday = new Declaration(new QualifiedMemberName(VbaModuleName, "vbThursday"), "VBA", "VbDayOfWeek", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbTuesday = new Declaration(new QualifiedMemberName(VbaModuleName, "vbTuesday"), "VBA", "VbDayOfWeek", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbUseSystemDayOfWeek = new Declaration(new QualifiedMemberName(VbaModuleName, "vbUseSystemDayOfWeek"), "VBA", "VbDayOfWeek", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbWednesday = new Declaration(new QualifiedMemberName(VbaModuleName, "vbWednesday"), "VBA", "VbDayOfWeek", true, false, Accessibility.Global, DeclarationType.EnumerationMember);

            public static Declaration VbFileAttribute = new Declaration(new QualifiedMemberName(VbaModuleName, "VbFileAttribute"), "VBA", "VbFileAttribute", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbNormal = new Declaration(new QualifiedMemberName(VbaModuleName, "vbNormal"), "VBA", "VbFileAttribute", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbReadOnly = new Declaration(new QualifiedMemberName(VbaModuleName, "vbReadOnly"), "VBA", "VbFileAttribute", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbHidden = new Declaration(new QualifiedMemberName(VbaModuleName, "vbHidden"), "VBA", "VbFileAttribute", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbSystem = new Declaration(new QualifiedMemberName(VbaModuleName, "vbSystem"), "VBA", "VbFileAttribute", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbVolume = new Declaration(new QualifiedMemberName(VbaModuleName, "vbVolume"), "VBA", "VbFileAttribute", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbDirectory = new Declaration(new QualifiedMemberName(VbaModuleName, "vbDirectory"), "VBA", "VbFileAttribute", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbArchive = new Declaration(new QualifiedMemberName(VbaModuleName, "vbArchive"), "VBA", "VbFileAttribute", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbAlias = new Declaration(new QualifiedMemberName(VbaModuleName, "vbAlias"), "VBA", "VbFileAttribute", true, false, Accessibility.Global, DeclarationType.EnumerationMember);

            public static Declaration VbFirstWeekOfYear = new Declaration(new QualifiedMemberName(VbaModuleName, "VbFirstWeekOfYear"), "VBA", "VbFirstWeekOfYear", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbFirstFourDays = new Declaration(new QualifiedMemberName(VbaModuleName, "vbFirstFourDays"), "VBA", "VbFirstWeekOfYear", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbFirstFullWeek = new Declaration(new QualifiedMemberName(VbaModuleName, "vbFirstFullWeek"), "VBA", "VbFirstWeekOfYear", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbFirstJan1 = new Declaration(new QualifiedMemberName(VbaModuleName, "vbFirstJan1"), "VBA", "VbFirstWeekOfYear", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbUseSystem = new Declaration(new QualifiedMemberName(VbaModuleName, "vbUseSystem"), "VBA", "VbFirstWeekOfYear", true, false, Accessibility.Global, DeclarationType.EnumerationMember);

            public static Declaration VbIMEStatus = new Declaration(new QualifiedMemberName(VbaModuleName, "VbIMEStatus"), "VBA", "VbIMEStatus", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbIMEAlphaDbl = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEAlphaDbl"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMEAlphaSng = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEAlphaSng"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMEDisable = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEDisable"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMEHiragana = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEHiragana"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMEKatakanaDbl = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEKatakanaDbl"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMEKatakanaSng = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEKatakanaSng"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMEModeAlpha = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeAlpha"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMEModeAlphaFull = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeAlphaFull"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMEModeDisable = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeDisable"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMEModeHangul = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeHangul"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMEModeHangulFull = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeHangulFull"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMEModeHiragana = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeHiragana"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMEModeKatakana = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeKatakana"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMEModeKatakanaHalf = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeKatakanaHalf"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMEModeNoControl = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeNoControl"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMEModeOff = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeOff"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMEModeOn = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeOn"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMENoOp = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMENoOp"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMEOff = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEOff"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIMEOn = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIMEOn"), "VBA", "VbIMEStatus", true, false, Accessibility.Global, DeclarationType.EnumerationMember);

            public static Declaration VbMsgBoxResult = new Declaration(new QualifiedMemberName(VbaModuleName, "VbMsgBoxResult"), "VBA", "VbMsgBoxResult", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbAbort = new Declaration(new QualifiedMemberName(VbaModuleName, "vbAbort"), "VBA", "VbMsgBoxResult", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbCancel = new Declaration(new QualifiedMemberName(VbaModuleName, "vbCancel"), "VBA", "VbMsgBoxResult", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbIgnore = new Declaration(new QualifiedMemberName(VbaModuleName, "vbIgnore"), "VBA", "VbMsgBoxResult", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbNo = new Declaration(new QualifiedMemberName(VbaModuleName, "vbNo"), "VBA", "VbMsgBoxResult", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbOk = new Declaration(new QualifiedMemberName(VbaModuleName, "vbOk"), "VBA", "VbMsgBoxResult", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbRetry = new Declaration(new QualifiedMemberName(VbaModuleName, "vbRetry"), "VBA", "VbMsgBoxResult", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbYes = new Declaration(new QualifiedMemberName(VbaModuleName, "vbYes"), "VBA", "VbMsgBoxResult", true, false, Accessibility.Global, DeclarationType.EnumerationMember);

            public static Declaration VbMsgBoxStyle = new Declaration(new QualifiedMemberName(VbaModuleName, "VbMsgBoxStyle"), "VBA", "VbMsgBoxStyle", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbAbortRetryIgnore = new Declaration(new QualifiedMemberName(VbaModuleName, "vbAbortRetryIgnore"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbApplicationModal = new Declaration(new QualifiedMemberName(VbaModuleName, "vbApplicationModal"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbCritical = new Declaration(new QualifiedMemberName(VbaModuleName, "vbCritical"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbDefaultButton1 = new Declaration(new QualifiedMemberName(VbaModuleName, "vbDefaultButton1"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbDefaultButton2 = new Declaration(new QualifiedMemberName(VbaModuleName, "vbDefaultButton2"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbDefaultButton3 = new Declaration(new QualifiedMemberName(VbaModuleName, "vbDefaultButton3"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbDefaultButton4 = new Declaration(new QualifiedMemberName(VbaModuleName, "vbDefaultButton4"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbExclamation = new Declaration(new QualifiedMemberName(VbaModuleName, "vbExclamation"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbInformation = new Declaration(new QualifiedMemberName(VbaModuleName, "vbInformation"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbMsgBoxHelpButton = new Declaration(new QualifiedMemberName(VbaModuleName, "vbMsgBoxHelpButton"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbMsgBoxRight = new Declaration(new QualifiedMemberName(VbaModuleName, "vbMsgBoxRight"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbMsgBoxRtlReading = new Declaration(new QualifiedMemberName(VbaModuleName, "vbMsgBoxRtlReading"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbMsgBoxSetForeground = new Declaration(new QualifiedMemberName(VbaModuleName, "vbMsgBoxSetForeground"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbOkCancel = new Declaration(new QualifiedMemberName(VbaModuleName, "vbOkCancel"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbOkOnly = new Declaration(new QualifiedMemberName(VbaModuleName, "vbOkOnly"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbQuestion = new Declaration(new QualifiedMemberName(VbaModuleName, "vbQuestion"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbRetryCancel = new Declaration(new QualifiedMemberName(VbaModuleName, "vbRetryCancel"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbSystemModal = new Declaration(new QualifiedMemberName(VbaModuleName, "vbSystemModal"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbYesNo = new Declaration(new QualifiedMemberName(VbaModuleName, "vbYesNo"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbYesNoCancel = new Declaration(new QualifiedMemberName(VbaModuleName, "vbYesNoCancel"), "VBA", "VbMsgBoxStyle", true, false, Accessibility.Global, DeclarationType.EnumerationMember);

            public static Declaration VbQueryClose = new Declaration(new QualifiedMemberName(VbaModuleName, "VbQueryClose"), "VBA", "VbQueryClose", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbAppTaskManager = new Declaration(new QualifiedMemberName(VbaModuleName, "vbAppTaskManager"), "VBA", "VbQueryClose", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbAppWindows = new Declaration(new QualifiedMemberName(VbaModuleName, "vbAppWindows"), "VBA", "VbQueryClose", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbFormCode = new Declaration(new QualifiedMemberName(VbaModuleName, "vbFormCode"), "VBA", "VbQueryClose", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbFormControlMenu = new Declaration(new QualifiedMemberName(VbaModuleName, "vbFormControlMenu"), "VBA", "VbQueryClose", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbFormMDIForm = new Declaration(new QualifiedMemberName(VbaModuleName, "vbFormMDIForm"), "VBA", "VbQueryClose", true, false, Accessibility.Global, DeclarationType.EnumerationMember);

            public static Declaration VbStrConv = new Declaration(new QualifiedMemberName(VbaModuleName, "VbStrConv"), "VBA", "VbStrConv", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbFromUnicode = new Declaration(new QualifiedMemberName(VbaModuleName, "vbFromUnicode"), "VBA", "VbStrConv", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbHiragana = new Declaration(new QualifiedMemberName(VbaModuleName, "vbHiragana"), "VBA", "VbStrConv", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbKatakana = new Declaration(new QualifiedMemberName(VbaModuleName, "vbKatakana"), "VBA", "VbStrConv", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbLowerCase = new Declaration(new QualifiedMemberName(VbaModuleName, "vbLowerCase"), "VBA", "VbStrConv", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbNarrow = new Declaration(new QualifiedMemberName(VbaModuleName, "vbNarrow"), "VBA", "VbStrConv", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbProperCase = new Declaration(new QualifiedMemberName(VbaModuleName, "vbProperCase"), "VBA", "VbStrConv", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbUnicode = new Declaration(new QualifiedMemberName(VbaModuleName, "vbUnicode"), "VBA", "VbStrConv", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbUpperCase = new Declaration(new QualifiedMemberName(VbaModuleName, "vbUpperCase"), "VBA", "VbStrConv", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbWide = new Declaration(new QualifiedMemberName(VbaModuleName, "vbWide"), "VBA", "VbStrConv", true, false, Accessibility.Global, DeclarationType.EnumerationMember);

            public static Declaration VbTriState = new Declaration(new QualifiedMemberName(VbaModuleName, "VbTriState"), "VBA", "VbTriState", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbFalse = new Declaration(new QualifiedMemberName(VbaModuleName, "vbFalse"), "VBA", "VbTriState", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbTrue = new Declaration(new QualifiedMemberName(VbaModuleName, "vbTrue"), "VBA", "VbTriState", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbUseDefault = new Declaration(new QualifiedMemberName(VbaModuleName, "vbUseDefault"), "VBA", "VbTriState", true, false, Accessibility.Global, DeclarationType.EnumerationMember);

            public static Declaration VbVarType = new Declaration(new QualifiedMemberName(VbaModuleName, "VbVarType"), "VBA", "VbVarType", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbArray = new Declaration(new QualifiedMemberName(VbaModuleName, "vbArray"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbBoolean = new Declaration(new QualifiedMemberName(VbaModuleName, "vbBoolean"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbByte = new Declaration(new QualifiedMemberName(VbaModuleName, "vbByte"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbCurrency = new Declaration(new QualifiedMemberName(VbaModuleName, "vbCurrency"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbDataObject = new Declaration(new QualifiedMemberName(VbaModuleName, "vbDataObject"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbDate = new Declaration(new QualifiedMemberName(VbaModuleName, "vbDate"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbDecimal = new Declaration(new QualifiedMemberName(VbaModuleName, "vbDecimal"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbDouble = new Declaration(new QualifiedMemberName(VbaModuleName, "vbDouble"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbEmpty = new Declaration(new QualifiedMemberName(VbaModuleName, "vbEmpty"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbError = new Declaration(new QualifiedMemberName(VbaModuleName, "vbError"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbInteger = new Declaration(new QualifiedMemberName(VbaModuleName, "vbInteger"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbLong = new Declaration(new QualifiedMemberName(VbaModuleName, "vbLong"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbLongLong = new Declaration(new QualifiedMemberName(VbaModuleName, "vbLongLong"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbNull = new Declaration(new QualifiedMemberName(VbaModuleName, "vbNull"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbObject = new Declaration(new QualifiedMemberName(VbaModuleName, "vbObject"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbSingle = new Declaration(new QualifiedMemberName(VbaModuleName, "vbSingle"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbString = new Declaration(new QualifiedMemberName(VbaModuleName, "vbString"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbUserDefinedType = new Declaration(new QualifiedMemberName(VbaModuleName, "vbUserDefinedType"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
            public static Declaration VbVariant = new Declaration(new QualifiedMemberName(VbaModuleName, "vbVariant"), "VBA", "VbVarType", true, false, Accessibility.Global, DeclarationType.EnumerationMember);
        }

        #region Predefined standard/procedural modules

        private class ColorConstantsModule
        {
            private static QualifiedModuleName ColorConstantsModuleName = new QualifiedModuleName("VBA", "ColorConstants");
            public static Declaration ColorConstants = new Declaration(new QualifiedMemberName(ColorConstantsModuleName, "ColorConstants"), "VBA", "ColorConstants", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration VbBlack = new Declaration(new QualifiedMemberName(ColorConstantsModuleName, "vbBlack"), "VBA.ColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbBlue = new Declaration(new QualifiedMemberName(ColorConstantsModuleName, "vbBlue"), "VBA.ColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbCyan = new Declaration(new QualifiedMemberName(ColorConstantsModuleName, "vbCyan"), "VBA.ColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbGreen = new Declaration(new QualifiedMemberName(ColorConstantsModuleName, "vbGreen"), "VBA.ColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbMagenta = new Declaration(new QualifiedMemberName(ColorConstantsModuleName, "vbMagenta"), "VBA.ColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbRed = new Declaration(new QualifiedMemberName(ColorConstantsModuleName, "vbRed"), "VBA.ColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbWhite = new Declaration(new QualifiedMemberName(ColorConstantsModuleName, "vbWhite"), "VBA.ColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbYellow = new Declaration(new QualifiedMemberName(ColorConstantsModuleName, "vbYellow"), "VBA.ColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
        }

        private class ConstantsModule
        {
            private static QualifiedModuleName ConstantsModuleName = new QualifiedModuleName("VBA", "Constants");
            public static Declaration Constants = new Declaration(new QualifiedMemberName(ConstantsModuleName, "Constants"), "VBA", "Constants", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration VbBack = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbBack"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbCr = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbCr"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbCrLf = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbCrLf"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbFormFeed = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbFormFeed"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbLf = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbLf"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbNewLine = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbNewLine"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbNullChar = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbNullChar"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbTab = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbTab"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbVerticalTab = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbVerticalTab"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbNullString = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbNullString"), "VBA.Constants", "String", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbObjectError = new Declaration(new QualifiedMemberName(ConstantsModuleName, "vbObjectError"), "VBA.Constants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
        }

        private class ConversionModule
        {
            private static QualifiedModuleName ConversionModuleName = new QualifiedModuleName("VBA", "Conversion");
            public static Declaration Conversion = new Declaration(new QualifiedMemberName(ConversionModuleName, "Conversion"), "VBA", "Conversion", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration CBool = new Declaration(new QualifiedMemberName(ConversionModuleName, "CBool"), "VBA.Conversion", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CByte = new Declaration(new QualifiedMemberName(ConversionModuleName, "CByte"), "VBA.Conversion", "Byte", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CCur = new Declaration(new QualifiedMemberName(ConversionModuleName, "CCur"), "VBA.Conversion", "Currency", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CDate = new Declaration(new QualifiedMemberName(ConversionModuleName, "CDate"), "VBA.Conversion", "Date", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CVDate = new Declaration(new QualifiedMemberName(ConversionModuleName, "CVDate"), "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CDbl = new Declaration(new QualifiedMemberName(ConversionModuleName, "CDbl"), "VBA.Conversion", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CDec = new Declaration(new QualifiedMemberName(ConversionModuleName, "CDec"), "VBA.Conversion", "Decimal", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CInt = new Declaration(new QualifiedMemberName(ConversionModuleName, "CInt"), "VBA.Conversion", "Integer", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CLng = new Declaration(new QualifiedMemberName(ConversionModuleName, "CLng"), "VBA.Conversion", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CLngLng = new Declaration(new QualifiedMemberName(ConversionModuleName, "CLngLng"), "VBA.Conversion", "LongLong", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CLngPtr = new Declaration(new QualifiedMemberName(ConversionModuleName, "CLngPtr"), "VBA.Conversion", "LongPtr", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CSng = new Declaration(new QualifiedMemberName(ConversionModuleName, "CSng"), "VBA.Conversion", "Single", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CStr = new Declaration(new QualifiedMemberName(ConversionModuleName, "CStr"), "VBA.Conversion", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CVar = new Declaration(new QualifiedMemberName(ConversionModuleName, "CVar"), "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CVErr = new Declaration(new QualifiedMemberName(ConversionModuleName, "CVErr"), "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Error = new Declaration(new QualifiedMemberName(ConversionModuleName, "Error"), "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration ErrorStr = new Declaration(new QualifiedMemberName(ConversionModuleName, "Error$"), "VBA.Conversion", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Fix = new Declaration(new QualifiedMemberName(ConversionModuleName, "Fix"), "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Hex = new Declaration(new QualifiedMemberName(ConversionModuleName, "Hex"), "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration HexStr = new Declaration(new QualifiedMemberName(ConversionModuleName, "Hex$"), "VBA.Conversion", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Int = new Declaration(new QualifiedMemberName(ConversionModuleName, "Int"), "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Oct = new Declaration(new QualifiedMemberName(ConversionModuleName, "Oct"), "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration OctStr = new Declaration(new QualifiedMemberName(ConversionModuleName, "Oct$"), "VBA.Conversion", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Str = new Declaration(new QualifiedMemberName(ConversionModuleName, "Str"), "VBA.Conversion", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration StrStr = new Declaration(new QualifiedMemberName(ConversionModuleName, "Str$"), "VBA.Conversion", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Val = new Declaration(new QualifiedMemberName(ConversionModuleName, "Val"), "VBA.Conversion", "Double", false, false, Accessibility.Global, DeclarationType.Function);
        }

        private class DateTimeModule
        {
            private static QualifiedModuleName DateTimeModuleName = new QualifiedModuleName("VBA", "DateTime");
            // functions
            public static Declaration DateTime = new Declaration(new QualifiedMemberName(DateTimeModuleName, "DateTime"), "VBA", "DateTime", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration DateAdd = new Declaration(new QualifiedMemberName(DateTimeModuleName, "DateAdd"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration DateDiff = new Declaration(new QualifiedMemberName(DateTimeModuleName, "DateDiff"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration DatePart = new Declaration(new QualifiedMemberName(DateTimeModuleName, "DatePart"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration DateSerial = new Declaration(new QualifiedMemberName(DateTimeModuleName, "DateSerial"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration DateValue = new Declaration(new QualifiedMemberName(DateTimeModuleName, "DateValue"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Day = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Day"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Hour = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Hour"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Minute = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Minute"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Month = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Month"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Second = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Second"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration TimeSerial = new Declaration(new QualifiedMemberName(DateTimeModuleName, "TimeSerial"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration TimeValue = new Declaration(new QualifiedMemberName(DateTimeModuleName, "TimeValue"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration WeekDay = new Declaration(new QualifiedMemberName(DateTimeModuleName, "WeekDay"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Year = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Year"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            // properties
            public static Declaration Calendar = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Calendar"), "VBA.DateTime", "vbCalendar", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Date = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Date"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration DateStr = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Date$"), "VBA.DateTime", "String", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Now = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Now"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Time = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Time"), "VBA.DateTime", "Variant", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration TimeStr = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Time$"), "VBA.DateTime", "String", false, false, Accessibility.Global, DeclarationType.PropertyGet);
            public static Declaration Timer = new Declaration(new QualifiedMemberName(DateTimeModuleName, "Timer"), "VBA.DateTime", "Single", false, false, Accessibility.Global, DeclarationType.PropertyGet);
        }

        private class FileSystemModule
        {
            private static QualifiedModuleName FileSystemModuleName = new QualifiedModuleName("VBA", "FileSystem");
            // functions
            public static Declaration FileSystem = new Declaration(new QualifiedMemberName(FileSystemModuleName, "FileSystem"), "VBA", "FileSystem", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration CurDir = new Declaration(new QualifiedMemberName(FileSystemModuleName, "CurDir"), "VBA.FileSystem", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CurDirStr = new Declaration(new QualifiedMemberName(FileSystemModuleName, "CurDir$"), "VBA.FileSystem", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Dir = new Declaration(new QualifiedMemberName(FileSystemModuleName, "Dir"), "VBA.FileSystem", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration EOF = new Declaration(new QualifiedMemberName(FileSystemModuleName, "EOF"), "VBA.FileSystem", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FileAttr = new Declaration(new QualifiedMemberName(FileSystemModuleName, "FileAttr"), "VBA.FileSystem", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FileDateTime = new Declaration(new QualifiedMemberName(FileSystemModuleName, "FileDateTime"), "VBA.FileSystem", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FileLen = new Declaration(new QualifiedMemberName(FileSystemModuleName, "FileLen"), "VBA.FileSystem", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FreeFile = new Declaration(new QualifiedMemberName(FileSystemModuleName, "FreeFile"), "VBA.FileSystem", "Integer", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Loc = new Declaration(new QualifiedMemberName(FileSystemModuleName, "Loc"), "VBA.FileSystem", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration LOF = new Declaration(new QualifiedMemberName(FileSystemModuleName, "LOF"), "VBA.FileSystem", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Seek = new Declaration(new QualifiedMemberName(FileSystemModuleName, "Seek"), "VBA.FileSystem", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            // procedures
            public static Declaration ChDir = new Declaration(new QualifiedMemberName(FileSystemModuleName, "ChDir"), "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration ChDrive = new Declaration(new QualifiedMemberName(FileSystemModuleName, "ChDrive"), "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration FileCopy = new Declaration(new QualifiedMemberName(FileSystemModuleName, "FileCopy"), "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration Kill = new Declaration(new QualifiedMemberName(FileSystemModuleName, "Kill"), "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration MkDir = new Declaration(new QualifiedMemberName(FileSystemModuleName, "MkDir"), "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration RmDir = new Declaration(new QualifiedMemberName(FileSystemModuleName, "RmDir"), "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration SetAttr = new Declaration(new QualifiedMemberName(FileSystemModuleName, "SetAttr"), "VBA.FileSystem", null, false, false, Accessibility.Global, DeclarationType.Procedure);
        }

        private class FinancialModule
        {
            private static QualifiedModuleName FinancialModuleName = new QualifiedModuleName("VBA", "Financial");
            public static Declaration Financial = new Declaration(new QualifiedMemberName(FinancialModuleName, "Financial"), "VBA", "Financial", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration DDB = new Declaration(new QualifiedMemberName(FinancialModuleName, "DDB"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FV = new Declaration(new QualifiedMemberName(FinancialModuleName, "FV"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IPmt = new Declaration(new QualifiedMemberName(FinancialModuleName, "IPmt"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IRR = new Declaration(new QualifiedMemberName(FinancialModuleName, "IRR"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration MIRR = new Declaration(new QualifiedMemberName(FinancialModuleName, "MIRR"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration NPer = new Declaration(new QualifiedMemberName(FinancialModuleName, "NPer"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration NPV = new Declaration(new QualifiedMemberName(FinancialModuleName, "NPV"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Pmt = new Declaration(new QualifiedMemberName(FinancialModuleName, "Pmt"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration PPmt = new Declaration(new QualifiedMemberName(FinancialModuleName, "PPmt"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration PV = new Declaration(new QualifiedMemberName(FinancialModuleName, "PV"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Rate = new Declaration(new QualifiedMemberName(FinancialModuleName, "Rate"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration SLN = new Declaration(new QualifiedMemberName(FinancialModuleName, "SLN"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration SYD = new Declaration(new QualifiedMemberName(FinancialModuleName, "SYD"), "VBA.Financial", "Double", false, false, Accessibility.Global, DeclarationType.Function);
        }

        private class InformationModule
        {
            private static QualifiedModuleName InformationModuleName = new QualifiedModuleName("VBA", "Information");
            public static Declaration Information = new Declaration(new QualifiedMemberName(InformationModuleName, "Information"), "VBA", "Information", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration Err = new Declaration(new QualifiedMemberName(InformationModuleName, "Err"), "VBA.Information", "ErrObject", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IMEStatus = new Declaration(new QualifiedMemberName(InformationModuleName, "IMEStatus"), "VBA.Information", "vbIMEStatus", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IsArray = new Declaration(new QualifiedMemberName(InformationModuleName, "IsArray"), "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IsDate = new Declaration(new QualifiedMemberName(InformationModuleName, "IsDate"), "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IsEmpty = new Declaration(new QualifiedMemberName(InformationModuleName, "IsEmpty"), "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IsError = new Declaration(new QualifiedMemberName(InformationModuleName, "IsError"), "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IsMissing = new Declaration(new QualifiedMemberName(InformationModuleName, "IsMissing"), "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IsNull = new Declaration(new QualifiedMemberName(InformationModuleName, "IsNull"), "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IsNumeric = new Declaration(new QualifiedMemberName(InformationModuleName, "IsNumeric"), "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IsObject = new Declaration(new QualifiedMemberName(InformationModuleName, "IsObject"), "VBA.Information", "Boolean", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration QBColor = new Declaration(new QualifiedMemberName(InformationModuleName, "QBColor"), "VBA.Information", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration RGB = new Declaration(new QualifiedMemberName(InformationModuleName, "RGB"), "VBA.Information", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration TypeName = new Declaration(new QualifiedMemberName(InformationModuleName, "TypeName"), "VBA.Information", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration VarType = new Declaration(new QualifiedMemberName(InformationModuleName, "VarType"), "VBA.Information", "vbVarType", false, false, Accessibility.Global, DeclarationType.Function);
        }

        private class InteractionModule
        {
            private static QualifiedModuleName InteractionModuleName = new QualifiedModuleName("VBA", "Interaction");
            // functions
            public static Declaration Interaction = new Declaration(new QualifiedMemberName(InteractionModuleName, "Interaction"), "VBA", "Interaction", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration CallByName = new Declaration(new QualifiedMemberName(InteractionModuleName, "CallByName"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Choose = new Declaration(new QualifiedMemberName(InteractionModuleName, "Choose"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Command = new Declaration(new QualifiedMemberName(InteractionModuleName, "Command"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CommandStr = new Declaration(new QualifiedMemberName(InteractionModuleName, "Command$"), "VBA.Interaction", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration CreateObject = new Declaration(new QualifiedMemberName(InteractionModuleName, "CreateObject"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration DoEvents = new Declaration(new QualifiedMemberName(InteractionModuleName, "DoEvents"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Environ = new Declaration(new QualifiedMemberName(InteractionModuleName, "Environ"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration EnvironStr = new Declaration(new QualifiedMemberName(InteractionModuleName, "Environ$"), "VBA.Interaction", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration GetAllSettings = new Declaration(new QualifiedMemberName(InteractionModuleName, "GetAllSettings"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration GetAttr = new Declaration(new QualifiedMemberName(InteractionModuleName, "GetAttr"), "VBA.Interaction", "vbFileAttribute", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration GetObject = new Declaration(new QualifiedMemberName(InteractionModuleName, "GetObject"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration GetSetting = new Declaration(new QualifiedMemberName(InteractionModuleName, "GetSetting"), "VBA.Interaction", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration IIf = new Declaration(new QualifiedMemberName(InteractionModuleName, "IIf"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration InputBox = new Declaration(new QualifiedMemberName(InteractionModuleName, "InputBox"), "VBA.Interaction", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration MsgBox = new Declaration(new QualifiedMemberName(InteractionModuleName, "MsgBox"), "VBA.Interaction", "vbMsgBoxResult", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Partition = new Declaration(new QualifiedMemberName(InteractionModuleName, "Partition"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Shell = new Declaration(new QualifiedMemberName(InteractionModuleName, "Shell"), "VBA.Interaction", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Switch = new Declaration(new QualifiedMemberName(InteractionModuleName, "Switch"), "VBA.Interaction", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            // procedures
            public static Declaration AppActivate = new Declaration(new QualifiedMemberName(InteractionModuleName, "AppActivate"), "VBA.Interaction", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration Beep = new Declaration(new QualifiedMemberName(InteractionModuleName, "Beep"), "VBA.Interaction", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration DeleteSetting = new Declaration(new QualifiedMemberName(InteractionModuleName, "DeleteSetting"), "VBA.Interaction", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration SaveSetting = new Declaration(new QualifiedMemberName(InteractionModuleName, "SaveSetting"), "VBA.Interaction", null, false, false, Accessibility.Global, DeclarationType.Procedure);
            public static Declaration SendKeys = new Declaration(new QualifiedMemberName(InteractionModuleName, "SendKeys"), "VBA.Interaction", null, false, false, Accessibility.Global, DeclarationType.Procedure);
        }

        private class KeyCodeConstantsModule
        {
            private static QualifiedModuleName KeyCodeConstantsModuleName = new QualifiedModuleName("VBA", "KeyCodeConstants");
            public static Declaration KeyCodeConstants = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "KeyCodeConstants"), "VBA", "KeyCodeConstants", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration VbKeyLButton = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyLButton"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyRButton = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyRButton"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyCancel = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyCancel"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyMButton = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyMButton"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyBack = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyBack"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyTab = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyTab"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyClear = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyClear"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyReturn = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyReturn"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyShift = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyShift"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyControl = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyControl"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyMenu = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyMenu"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyPause = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyPause"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyCapital = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyCapital"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyEscape = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyEscape"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeySpace = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeySpace"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyPageUp = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyPageUp"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyPageDown = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyPageDown"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyEnd = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyEnd"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyHome = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyHome"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyLeft = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyLeft"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyUp = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyUp"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyRight = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyRight"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyDown = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyDown"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeySelect = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeySelect"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyPrint = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyPrint"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyExecute = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyExecute"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeySnapshot = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeySnapshot"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyInsert = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyInsert"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyDelete = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyDelete"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyHelp = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyHelp"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyNumLock = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumLock"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyA = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyA"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyB = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyB"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyC = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyC"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyD = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyD"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyE = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyE"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyF = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyG = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyG"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyH = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyH"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyI = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyI"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyJ = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyJ"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyK = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyK"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyL = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyL"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyM = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyM"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyN = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyN"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyO = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyO"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyP = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyP"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyQ = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyQ"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyR = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyR"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyS = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyS"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyT = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyT"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyU = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyU"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyV = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyV"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyW = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyW"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyX = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyX"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyY = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyY"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyZ = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyZ"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKey0 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey0"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKey1 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey1"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKey2 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey2"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKey3 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey3"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKey4 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey4"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKey5 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey5"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKey6 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey6"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKey7 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey7"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKey8 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey8"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKey9 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKey9"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyNumpad0 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad0"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyNumpad1 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad1"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyNumpad2 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad2"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyNumpad3 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad3"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyNumpad4 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad4"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyNumpad5 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad5"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyNumpad6 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad6"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyNumpad7 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad7"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyNumpad8 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad8"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyNumpad9 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyNumpad9"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyMultiply = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyMultiply"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyAdd = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyAdd"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeySeparator = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeySeparator"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeySubtract = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeySubtract"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyDecimal = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyDecimal"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyDivide = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyDivide"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyF1 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF1"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyF2 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF2"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyF3 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF3"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyF4 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF4"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyF5 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF5"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyF6 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF6"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyF7 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF7"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyF8 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF8"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyF9 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF9"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyF10 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF10"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyF11 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF11"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyF12 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF12"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyF13 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF13"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyF14 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF14"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyF15 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF15"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbKeyF16 = new Declaration(new QualifiedMemberName(KeyCodeConstantsModuleName, "vbKeyF16"), "VBA.KeyCodeConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
        }

        private class MathModule
        {
            private static QualifiedModuleName MathModuleName = new QualifiedModuleName("VBA", "Math");
            // functions
            public static Declaration Math = new Declaration(new QualifiedMemberName(MathModuleName, "Math"), "VBA", "Math", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration Abs = new Declaration(new QualifiedMemberName(MathModuleName, "Abs"), "VBA.Math", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Atn = new Declaration(new QualifiedMemberName(MathModuleName, "Atn"), "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Cos = new Declaration(new QualifiedMemberName(MathModuleName, "Cos"), "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Exp = new Declaration(new QualifiedMemberName(MathModuleName, "Exp"), "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Log = new Declaration(new QualifiedMemberName(MathModuleName, "Log"), "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Rnd = new Declaration(new QualifiedMemberName(MathModuleName, "Rnd"), "VBA.Math", "Single", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Round = new Declaration(new QualifiedMemberName(MathModuleName, "Round"), "VBA.Math", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Sgn = new Declaration(new QualifiedMemberName(MathModuleName, "Sgn"), "VBA.Math", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Sin = new Declaration(new QualifiedMemberName(MathModuleName, "Sin"), "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Sqr = new Declaration(new QualifiedMemberName(MathModuleName, "Sqr"), "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Tan = new Declaration(new QualifiedMemberName(MathModuleName, "Tan"), "VBA.Math", "Double", false, false, Accessibility.Global, DeclarationType.Function);
            //procedures
            public static Declaration Randomize = new Declaration(new QualifiedMemberName(MathModuleName, "Randomize"), "VBA.Math", null, false, false, Accessibility.Global, DeclarationType.Procedure);
        }

        private class StringsModule
        {
            private static QualifiedModuleName StringsModuleName = new QualifiedModuleName("VBA", "Strings");
            public static Declaration Strings = new Declaration(new QualifiedMemberName(StringsModuleName, "Strings"), "VBA", "Strings", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration Asc = new Declaration(new QualifiedMemberName(StringsModuleName, "Asc"), "VBA.Strings", "Integer", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration AscW = new Declaration(new QualifiedMemberName(StringsModuleName, "AscW"), "VBA.Strings", "Integer", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration AscB = new Declaration(new QualifiedMemberName(StringsModuleName, "AscB"), "VBA.Strings", "Integer", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Chr = new Declaration(new QualifiedMemberName(StringsModuleName, "Chr"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration ChrStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Chr$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration ChrB = new Declaration(new QualifiedMemberName(StringsModuleName, "ChrB"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration ChrBStr = new Declaration(new QualifiedMemberName(StringsModuleName, "ChrB$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration ChrW = new Declaration(new QualifiedMemberName(StringsModuleName, "ChrW"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration ChrWStr = new Declaration(new QualifiedMemberName(StringsModuleName, "ChrW$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Filter = new Declaration(new QualifiedMemberName(StringsModuleName, "Filter"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Format = new Declaration(new QualifiedMemberName(StringsModuleName, "Format"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FormatStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Format$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FormatCurrency = new Declaration(new QualifiedMemberName(StringsModuleName, "FormatCurrency"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FormatDateTime = new Declaration(new QualifiedMemberName(StringsModuleName, "FormatDateTime"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FormatNumber = new Declaration(new QualifiedMemberName(StringsModuleName, "FormatNumber"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration FormatPercent = new Declaration(new QualifiedMemberName(StringsModuleName, "FormatPercent"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration InStr = new Declaration(new QualifiedMemberName(StringsModuleName, "InStr"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration InStrB = new Declaration(new QualifiedMemberName(StringsModuleName, "InStrB"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration InStrRev = new Declaration(new QualifiedMemberName(StringsModuleName, "InStrRev"), "VBA.Strings", "Long", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Join = new Declaration(new QualifiedMemberName(StringsModuleName, "Join"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration LCase = new Declaration(new QualifiedMemberName(StringsModuleName, "LCase"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration LCaseStr = new Declaration(new QualifiedMemberName(StringsModuleName, "LCase$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Left = new Declaration(new QualifiedMemberName(StringsModuleName, "Left"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration LeftB = new Declaration(new QualifiedMemberName(StringsModuleName, "LeftB"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration LeftStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Left$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration LeftBStr = new Declaration(new QualifiedMemberName(StringsModuleName, "LeftB$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Len = new Declaration(new QualifiedMemberName(StringsModuleName, "Len"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration LenB = new Declaration(new QualifiedMemberName(StringsModuleName, "LenB"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration LTrim = new Declaration(new QualifiedMemberName(StringsModuleName, "LTrim"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration RTrim = new Declaration(new QualifiedMemberName(StringsModuleName, "RTrim"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Trim = new Declaration(new QualifiedMemberName(StringsModuleName, "Trim"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration LTrimStr = new Declaration(new QualifiedMemberName(StringsModuleName, "LTrim$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration RTrimStr = new Declaration(new QualifiedMemberName(StringsModuleName, "RTrim$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration TrimStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Trim$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Mid = new Declaration(new QualifiedMemberName(StringsModuleName, "Mid"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration MidB = new Declaration(new QualifiedMemberName(StringsModuleName, "MidB"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration MidStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Mid$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration MidBStr = new Declaration(new QualifiedMemberName(StringsModuleName, "MidB$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration MonthName = new Declaration(new QualifiedMemberName(StringsModuleName, "MonthName"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Replace = new Declaration(new QualifiedMemberName(StringsModuleName, "Replace"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Right = new Declaration(new QualifiedMemberName(StringsModuleName, "Right"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration RightB = new Declaration(new QualifiedMemberName(StringsModuleName, "RightB"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration RightStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Right$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration RightBStr = new Declaration(new QualifiedMemberName(StringsModuleName, "RightB$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Space = new Declaration(new QualifiedMemberName(StringsModuleName, "Space"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration SpaceStr = new Declaration(new QualifiedMemberName(StringsModuleName, "Space$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Split = new Declaration(new QualifiedMemberName(StringsModuleName, "Split"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration StrComp = new Declaration(new QualifiedMemberName(StringsModuleName, "StrComp"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration StrConv = new Declaration(new QualifiedMemberName(StringsModuleName, "StrConv"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration String = new Declaration(new QualifiedMemberName(StringsModuleName, "String"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration StringStr = new Declaration(new QualifiedMemberName(StringsModuleName, "String$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration StrReverse = new Declaration(new QualifiedMemberName(StringsModuleName, "StrReverse"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration UCase = new Declaration(new QualifiedMemberName(StringsModuleName, "UCase"), "VBA.Strings", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration UCaseStr = new Declaration(new QualifiedMemberName(StringsModuleName, "UCase$"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration WeekdayName = new Declaration(new QualifiedMemberName(StringsModuleName, "WeekdayName"), "VBA.Strings", "String", false, false, Accessibility.Global, DeclarationType.Function);
        }

        private class SystemColorConstantsModule
        {
            private static QualifiedModuleName SystemColorConstantsModuleName = new QualifiedModuleName("VBA", "SystemColorConstants");
            public static Declaration SystemColorConstants = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "SystemColorConstants"), "VBA", "SystemColorConstants", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration VbScrollBars = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbScrollBars"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbDesktop = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbDesktop"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbActiveTitleBar = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbActiveTitleBar"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbInactiveTitleBar = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbInactiveTitleBar"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbMenuBar = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbMenuBar"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbWindowBackground = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbWindowBackground"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbWindowFrame = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbWindowFrame"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbMenuText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbMenuText"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbWindowText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbWindowText"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbTitleBarText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbTitleBarText"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbActiveBorder = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbActiveBorder"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbInactiveBorder = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbInactiveBorder"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbApplicationWorkspace = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbApplicationWorkspace"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbHighlight = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbHighlight"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbHighlightText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbHighlightText"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbButtonFace = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbButtonFace"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbButtonShadow = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbButtonShadow"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbGrayText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbGrayText"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbButtonText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbButtonText"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbInactiveCaptionText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbInactiveCaptionText"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration Vb3DHighlight = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vb3DHighlight"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration Vb3DDKShadow = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vb3DDKShadow"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration Vb3DLight = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vb3DLight"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration Vb3DFace = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vb3DFace"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration Vb3DShadow = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vb3DShadow"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbInfoText = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbInfoText"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
            public static Declaration VbInfoBackground = new Declaration(new QualifiedMemberName(SystemColorConstantsModuleName, "vbInfoBackground"), "VBA.SystemColorConstants", "Long", true, false, Accessibility.Global, DeclarationType.Constant);
        }

        #endregion

        #region Predefined class modules

        private class CollectionClass
        {
            public static Declaration Collection = new Declaration(new QualifiedMemberName(VbaModuleName, "Collection"), "VBA", "Collection", false, false, Accessibility.Global, DeclarationType.Class);
            public static Declaration Count = new Declaration(new QualifiedMemberName(VbaModuleName, "Count"), "VBA.Collection", "Long", false, false, Accessibility.Public, DeclarationType.Function);
            public static Declaration Item = new Declaration(new QualifiedMemberName(VbaModuleName, "Item"), "VBA.Collection", "Variant", false, false, Accessibility.Public, DeclarationType.Function);
            public static Declaration Add = new Declaration(new QualifiedMemberName(VbaModuleName, "Add"), "VBA.Collection", null, false, false, Accessibility.Public, DeclarationType.Procedure);
            public static Declaration Remove = new Declaration(new QualifiedMemberName(VbaModuleName, "Remove"), "VBA.Collection", null, false, false, Accessibility.Public, DeclarationType.Procedure);
        }

        private class ErrObjectClass
        {
            public static Declaration ErrObject = new Declaration(new QualifiedMemberName(VbaModuleName, "ErrObject"), "VBA", "ErrObject", false, false, Accessibility.Global, DeclarationType.Class);
            public static Declaration Clear = new Declaration(new QualifiedMemberName(VbaModuleName, "Clear"), "VBA.ErrObject", null, false, false, Accessibility.Public, DeclarationType.Procedure);
            public static Declaration Raise = new Declaration(new QualifiedMemberName(VbaModuleName, "Raise"), "VBA.ErrObject", null, false, false, Accessibility.Public, DeclarationType.Procedure);
            public static Declaration Description = new Declaration(new QualifiedMemberName(VbaModuleName, "Description"), "VBA.ErrObject", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration HelpContext = new Declaration(new QualifiedMemberName(VbaModuleName, "HelpContext"), "VBA.ErrObject", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration HelpFile = new Declaration(new QualifiedMemberName(VbaModuleName, "HelpFile"), "VBA.ErrObject", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration LastDllError = new Declaration(new QualifiedMemberName(VbaModuleName, "LastDllError"), "VBA.ErrObject", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration Number = new Declaration(new QualifiedMemberName(VbaModuleName, "Number"), "VBA.ErrObject", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static Declaration Source = new Declaration(new QualifiedMemberName(VbaModuleName, "Source"), "VBA.ErrObject", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet);
        }

        private class GlobalClass
        {
            public static Declaration Global = new Declaration(new QualifiedMemberName(VbaModuleName, "Global"), "VBA", "Global", false, false, Accessibility.Global, DeclarationType.Class);
            public static Declaration Load = new Declaration(new QualifiedMemberName(VbaModuleName, "Load"), "VBA.Global", null, false, false, Accessibility.Public, DeclarationType.Procedure);
            public static Declaration Unload = new Declaration(new QualifiedMemberName(VbaModuleName, "Unload"), "VBA.Global", null, false, false, Accessibility.Public, DeclarationType.Procedure);
        }
        
        #endregion

    }
}