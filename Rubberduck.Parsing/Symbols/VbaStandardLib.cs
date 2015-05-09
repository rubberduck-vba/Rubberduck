using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Reflection;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    /// <summary>
    /// Defines <see cref="Declaration"/> objects for the standard library.
    /// </summary>
    internal static class VbaStandardLib
    {
        private static IEnumerable<Declaration> _standardLibDeclarations;
        private static readonly QualifiedModuleName VbaModuleName = new QualifiedModuleName("VBA", "VBA");

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
            public static Declaration Vba = new Declaration(new QualifiedMemberName(VbaModuleName, "VBA"), "VBA", "VBA", true, false, Accessibility.Global, DeclarationType.Project);

            public static Declaration FormShowConstants = new Declaration(new QualifiedMemberName(VbaModuleName, "FormShowConstants"), "VBA", "FormShowConstants", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbModal = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbModal"), "VBA", "FormShowConstants", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbModeless = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbModeless"), "VBA", "FormShowConstants", Accessibility.Global, DeclarationType.EnumerationMember, "0");

            public static Declaration VbAppWinStyle = new Declaration(new QualifiedMemberName(VbaModuleName, "VbAppWinStyle"), "VBA", "VbAppWinStyle", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbHide = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbHide"), "VBA", "VbAppWinStyle", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbMaximizedFocus = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMaximizedFocus"), "VBA", "VbAppWinStyle", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbMinimizedFocus = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMinimizedFocus"), "VBA", "VbAppWinStyle", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbMinimizedNoFocus = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMinimizedNoFocus"), "VBA", "VbAppWinStyle", Accessibility.Global, DeclarationType.EnumerationMember, "6");
            public static Declaration VbNormalFocus = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbNormalFocus"), "VBA", "VbAppWinStyle", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbNormalNoFocus = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbNormalNoFocus"), "VBA", "VbAppWinStyle", Accessibility.Global, DeclarationType.EnumerationMember, "4");

            public static Declaration VbCalendar = new Declaration(new QualifiedMemberName(VbaModuleName, "VbCalendar"), "VBA", "VbCalendar", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbCalGreg = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbCalGreg"), "VBA", "VbCalendar", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbCalHijri = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbCalHijri"), "VBA", "VbCalendar", Accessibility.Global, DeclarationType.EnumerationMember, "1");

            public static Declaration VbCallType = new Declaration(new QualifiedMemberName(VbaModuleName, "VbCallType"), "VBA", "VbCallType", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbGet = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbGet"), "VBA", "VbCallType", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbLet = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbLet"), "VBA", "VbCallType", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration VbMethod = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMethod"), "VBA", "VbCallType", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbSet = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbSet"), "VBA", "VbCallType", Accessibility.Global, DeclarationType.EnumerationMember, "8");

            public static Declaration VbCompareMethod = new Declaration(new QualifiedMemberName(VbaModuleName, "VbCompareMethod"), "VBA", "VbCompareMethod", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbBinaryCompare = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbBinaryCompare"), "VBA", "VbCompareMethod", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbTextCompare = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbTextCompare"), "VBA", "VbCompareMethod", Accessibility.Global, DeclarationType.EnumerationMember, "1");

            public static Declaration VbDateTimeFormat = new Declaration(new QualifiedMemberName(VbaModuleName, "VbDateTimeFormat"), "VBA", "VbDateTimeFormat", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbGeneralDate = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbGeneralDate"), "VBA", "VbDateTimeFormat", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbLongDate = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbLongDate"), "VBA", "VbDateTimeFormat", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbLongTime = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbLongTime"), "VBA", "VbDateTimeFormat", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbShortDate = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbShortDate"), "VBA", "VbDateTimeFormat", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbShortTime = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbShortTime"), "VBA", "VbDateTimeFormat", Accessibility.Global, DeclarationType.EnumerationMember, "4");

            public static Declaration VbDayOfWeek = new Declaration(new QualifiedMemberName(VbaModuleName, "VbDayOfWeek"), "VBA", "VbDayOfWeek", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbFriday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFriday"), "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "6");
            public static Declaration VbMonday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMonday"), "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbSaturday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbSaturday"), "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "7");
            public static Declaration VbSunday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbSunday"), "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbThursday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbThursday"), "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "5");
            public static Declaration VbTuesday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbTuesday"), "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbUseSystemDayOfWeek = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbUseSystemDayOfWeek"), "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbWednesday = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbWednesday"), "VBA", "VbDayOfWeek", Accessibility.Global, DeclarationType.EnumerationMember, "4");

            public static Declaration VbFileAttribute = new Declaration(new QualifiedMemberName(VbaModuleName, "VbFileAttribute"), "VBA", "VbFileAttribute", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbNormal = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbNormal"), "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbReadOnly = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbReadOnly"), "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbHidden = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbHidden"), "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbSystem = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbSystem"), "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration VbVolume = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbVolume"), "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "8");
            public static Declaration VbDirectory = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDirectory"), "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "16");
            public static Declaration VbArchive = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbArchive"), "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "32");
            public static Declaration VbAlias = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbAlias"), "VBA", "VbFileAttribute", Accessibility.Global, DeclarationType.EnumerationMember, "64");

            public static Declaration VbFirstWeekOfYear = new Declaration(new QualifiedMemberName(VbaModuleName, "VbFirstWeekOfYear"), "VBA", "VbFirstWeekOfYear", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbFirstFourDays = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFirstFourDays"), "VBA", "VbFirstWeekOfYear", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbFirstFullWeek = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFirstFullWeek"), "VBA", "VbFirstWeekOfYear", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbFirstJan1 = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFirstJan1"), "VBA", "VbFirstWeekOfYear", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbUseSystem = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbUseSystem"), "VBA", "VbFirstWeekOfYear", Accessibility.Global, DeclarationType.EnumerationMember, "0");

            public static Declaration VbIMEStatus = new Declaration(new QualifiedMemberName(VbaModuleName, "VbIMEStatus"), "VBA", "VbIMEStatus", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbIMEAlphaDbl = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEAlphaDbl"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "7");
            public static Declaration VbIMEAlphaSng = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEAlphaSng"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "8");
            public static Declaration VbIMEDisable = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEDisable"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbIMEHiragana = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEHiragana"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration VbIMEKatakanaDbl = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEKatakanaDbl"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "5");
            public static Declaration VbIMEKatakanaSng = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEKatakanaSng"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "6");
            public static Declaration VbIMEModeAlpha = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeAlpha"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "8");
            public static Declaration VbIMEModeAlphaFull = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeAlphaFull"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "7");
            public static Declaration VbIMEModeDisable = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeDisable"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbIMEModeHangul = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeHangul"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "10");
            public static Declaration VbIMEModeHangulFull = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeHangulFull"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "9");
            public static Declaration VbIMEModeHiragana = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeHiragana"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration VbIMEModeKatakana = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeKatakana"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "5");
            public static Declaration VbIMEModeKatakanaHalf = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeKatakanaHalf"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "6");
            public static Declaration VbIMEModeNoControl = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeNoControl"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember,"0");
            public static Declaration VbIMEModeOff = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeOff"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbIMEModeOn = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEModeOn"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbIMENoOp = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMENoOp"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbIMEOff = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEOff"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbIMEOn = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIMEOn"), "VBA", "VbIMEStatus", Accessibility.Global, DeclarationType.EnumerationMember, "1");

            public static Declaration VbMsgBoxResult = new Declaration(new QualifiedMemberName(VbaModuleName, "VbMsgBoxResult"), "VBA", "VbMsgBoxResult", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbAbort = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbAbort"), "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbCancel = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbCancel"), "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbIgnore = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbIgnore"), "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "5");
            public static Declaration VbNo = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbNo"), "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "7");
            public static Declaration VbOk = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbOk"), "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbRetry = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbRetry"), "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration VbYes = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbYes"), "VBA", "VbMsgBoxResult", Accessibility.Global, DeclarationType.EnumerationMember, "6");

            public static Declaration VbMsgBoxStyle = new Declaration(new QualifiedMemberName(VbaModuleName, "VbMsgBoxStyle"), "VBA", "VbMsgBoxStyle", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbAbortRetryIgnore = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbAbortRetryIgnore"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbApplicationModal = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbApplicationModal"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbCritical = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbCritical"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "16");
            public static Declaration VbDefaultButton1 = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDefaultButton1"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbDefaultButton2 = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDefaultButton2"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "256");
            public static Declaration VbDefaultButton3 = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDefaultButton3"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "512");
            public static Declaration VbDefaultButton4 = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDefaultButton4"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "768");
            public static Declaration VbExclamation = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbExclamation"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "48");
            public static Declaration VbInformation = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbInformation"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "64");
            public static Declaration VbMsgBoxHelpButton = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMsgBoxHelpButton"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "16384");
            public static Declaration VbMsgBoxRight = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMsgBoxRight"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "524288");
            public static Declaration VbMsgBoxRtlReading = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMsgBoxRtlReading"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "1048576");
            public static Declaration VbMsgBoxSetForeground = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbMsgBoxSetForeground"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "65536");
            public static Declaration VbOkCancel = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbOkCancel"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbOkOnly = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbOkOnly"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbQuestion = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbQuestion"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "32");
            public static Declaration VbRetryCancel = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbRetryCancel"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "5");
            public static Declaration VbSystemModal = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbSystemModal"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "4096");
            public static Declaration VbYesNo = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbYesNo"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration VbYesNoCancel = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbYesNoCancel"), "VBA", "VbMsgBoxStyle", Accessibility.Global, DeclarationType.EnumerationMember, "3");

            public static Declaration VbQueryClose = new Declaration(new QualifiedMemberName(VbaModuleName, "VbQueryClose"), "VBA", "VbQueryClose", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbAppTaskManager = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbAppTaskManager"), "VBA", "VbQueryClose", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbAppWindows = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbAppWindows"), "VBA", "VbQueryClose", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbFormCode = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFormCode"), "VBA", "VbQueryClose", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbFormControlMenu = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFormControlMenu"), "VBA", "VbQueryClose", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbFormMDIForm = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFormMDIForm"), "VBA", "VbQueryClose", Accessibility.Global, DeclarationType.EnumerationMember, "4");

            public static Declaration VbStrConv = new Declaration(new QualifiedMemberName(VbaModuleName, "VbStrConv"), "VBA", "VbStrConv", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbFromUnicode = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFromUnicode"), "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "128");
            public static Declaration VbHiragana = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbHiragana"), "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "32");
            public static Declaration VbKatakana = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbKatakana"), "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "16");
            public static Declaration VbLowerCase = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbLowerCase"), "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbNarrow = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbNarrow"), "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "8");
            public static Declaration VbProperCase = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbProperCase"), "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbUnicode = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbUnicode"), "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "64");
            public static Declaration VbUpperCase = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbUpperCase"), "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbWide = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbWide"), "VBA", "VbStrConv", Accessibility.Global, DeclarationType.EnumerationMember, "4");

            public static Declaration VbTriState = new Declaration(new QualifiedMemberName(VbaModuleName, "VbTriState"), "VBA", "VbTriState", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbFalse = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbFalse"), "VBA", "VbTriState", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbTrue = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbTrue"), "VBA", "VbTriState", Accessibility.Global, DeclarationType.EnumerationMember, "-1");
            public static Declaration VbUseDefault = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbUseDefault"), "VBA", "VbTriState", Accessibility.Global, DeclarationType.EnumerationMember, "-2");

            public static Declaration VbVarType = new Declaration(new QualifiedMemberName(VbaModuleName, "VbVarType"), "VBA", "VbVarType", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration VbArray = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbArray"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "8192");
            public static Declaration VbBoolean = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbBoolean"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "11");
            public static Declaration VbByte = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbByte"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "17");
            public static Declaration VbCurrency = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbCurrency"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "6");
            public static Declaration VbDataObject = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDataObject"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "13");
            public static Declaration VbDate = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDate"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "7");
            public static Declaration VbDecimal = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDecimal"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "14");
            public static Declaration VbDouble = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbDouble"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "5");
            public static Declaration VbEmpty = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbEmpty"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration VbError = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbError"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "10");
            public static Declaration VbInteger = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbInteger"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration VbLong = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbLong"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration VbLongLong = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbLongLong"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "20");
            public static Declaration VbNull = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbNull"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration VbObject = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbObject"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "9");
            public static Declaration VbSingle = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbSingle"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration VbString = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbString"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "8");
            public static Declaration VbUserDefinedType = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbUserDefinedType"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "36");
            public static Declaration VbVariant = new ValuedDeclaration(new QualifiedMemberName(VbaModuleName, "vbVariant"), "VBA", "VbVarType", Accessibility.Global, DeclarationType.EnumerationMember, "12");
        }

        #region Predefined standard/procedural modules

        private class ColorConstantsModule
        {
            private static QualifiedModuleName ColorConstantsModuleName = new QualifiedModuleName("VBA", "ColorConstants");
            public static Declaration ColorConstants = new Declaration(new QualifiedMemberName(ColorConstantsModuleName, "ColorConstants"), "VBA", "ColorConstants", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration VbBlack = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbBlack"), "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "0");
            public static Declaration VbBlue = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbBlue"), "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "16711680");
            public static Declaration VbCyan = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbCyan"), "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "16776960");
            public static Declaration VbGreen = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbGreen"), "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "65280");
            public static Declaration VbMagenta = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbMagenta"), "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "16711935");
            public static Declaration VbRed = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbRed"), "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "255");
            public static Declaration VbWhite = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbWhite"), "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "16777215");
            public static Declaration VbYellow = new ValuedDeclaration(new QualifiedMemberName(ColorConstantsModuleName, "vbYellow"), "VBA.ColorConstants", "Long", Accessibility.Global, DeclarationType.Constant, "65535");
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
            public static Declaration VbObjectError = new ValuedDeclaration(new QualifiedMemberName(ConstantsModuleName, "vbObjectError"), "VBA.Constants", "Long", Accessibility.Global, DeclarationType.Constant, "-2147221504");
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
            public static Declaration MacID = new Declaration(new QualifiedMemberName(ConversionModuleName, "MacID"), "VBA.Conversion", "Long", false, false, Accessibility.Global, DeclarationType.Function);
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

        private class HiddenModule
        {
            private static QualifiedModuleName HiddenModuleName = new QualifiedModuleName("VBA", "[_HiddenModule]");
            public static Declaration Array = new Declaration(new QualifiedMemberName(HiddenModuleName, "Array"), "VBA.[_HiddenModule]", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Input = new Declaration(new QualifiedMemberName(HiddenModuleName, "Input"), "VBA.[_HiddenModule]", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration InputStr = new Declaration(new QualifiedMemberName(HiddenModuleName, "Input$"), "VBA.[_HiddenModule]", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration InputB = new Declaration(new QualifiedMemberName(HiddenModuleName, "InputB"), "VBA.[_HiddenModule]", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration InputBStr = new Declaration(new QualifiedMemberName(HiddenModuleName, "InputB$"), "VBA.[_HiddenModule]", "String", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Width = new Declaration(new QualifiedMemberName(HiddenModuleName, "Width"), "VBA.[_HiddenModule]", "Variant", false, false, Accessibility.Global, DeclarationType.Function);
            
            // hidden members... of hidden module (like, very very hidden!)
            public static Declaration ObjPtr = new Declaration(new QualifiedMemberName(HiddenModuleName, "ObjPtr"), "VBA.[_HiddenModule]", "LongPtr", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration StrPtr = new Declaration(new QualifiedMemberName(HiddenModuleName, "StrPtr"), "VBA.[_HiddenModule]", "LongPtr", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration VarPtr = new Declaration(new QualifiedMemberName(HiddenModuleName, "VarPtr"), "VBA.[_HiddenModule]", "LongPtr", false, false, Accessibility.Global, DeclarationType.Function);
        }

        private class InformationModule
        {
            private static QualifiedModuleName InformationModuleName = new QualifiedModuleName("VBA", "Information");
            public static Declaration Information = new Declaration(new QualifiedMemberName(InformationModuleName, "Information"), "VBA", "Information", false, false, Accessibility.Global, DeclarationType.Module);
            public static Declaration Err = new Declaration(new QualifiedMemberName(InformationModuleName, "Err"), "VBA.Information", "ErrObject", false, false, Accessibility.Global, DeclarationType.Function);
            public static Declaration Erl = new Declaration(new QualifiedMemberName(InformationModuleName, "Erl"), "VBA.Information", "Long", false, false, Accessibility.Global, DeclarationType.Function);
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
            public static Declaration MacScript = new Declaration(new QualifiedMemberName(InteractionModuleName, "MacScript"), "VBA.Interaction", "String", false, false, Accessibility.Global, DeclarationType.Function);
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
            //todo: define these constants as ValuedDeclaration items. values at https://msdn.microsoft.com/en-us/library/ee199372.aspx
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
            public static Declaration NewEnum = new Declaration(new QualifiedMemberName(VbaModuleName, "[_NewEnum]"), "VBA.Collection", "Unknown", false, false, Accessibility.Public, DeclarationType.Function);
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

            public static Declaration UserForms = new Declaration(new QualifiedMemberName(VbaModuleName, "UserForms"), "VBA.Global", "Object", true, false, Accessibility.Public, DeclarationType.PropertyGet);
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
            public static Declaration UserForm = new Declaration(new QualifiedMemberName(MsFormsModuleName, "UserForm"), "MSForms", "UserForm", true, false, Accessibility.Global, DeclarationType.Class);

            // events
            public static Declaration AddControl = new Declaration(new QualifiedMemberName(MsFormsModuleName, "AddControl"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration BeforeDragOver = new Declaration(new QualifiedMemberName(MsFormsModuleName, "BeforeDragOver"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration BeforeDropOrPaste = new Declaration(new QualifiedMemberName(MsFormsModuleName, "BeforeDropOrPaste"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Click = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Click"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration DblClick = new Declaration(new QualifiedMemberName(MsFormsModuleName, "DblClick"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Error = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Error"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration KeyDown = new Declaration(new QualifiedMemberName(MsFormsModuleName, "KeyDown"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration KeyPress = new Declaration(new QualifiedMemberName(MsFormsModuleName, "KeyPress"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration KeyUp = new Declaration(new QualifiedMemberName(MsFormsModuleName, "KeyUp"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Layout = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Layout"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration MouseDown = new Declaration(new QualifiedMemberName(MsFormsModuleName, "MouseDown"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration MouseMove = new Declaration(new QualifiedMemberName(MsFormsModuleName, "MouseMove"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration MouseUp = new Declaration(new QualifiedMemberName(MsFormsModuleName, "MouseUp"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration RemoveControl = new Declaration(new QualifiedMemberName(MsFormsModuleName, "RemoveControl"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Scroll = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Scroll"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Zoom = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Zoom"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);

            // ghost events (nowhere in the object browser)
            public static Declaration Activate = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Activate"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Deactivate = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Deactivate"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Initialize = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Initialize"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration QueryClose = new Declaration(new QualifiedMemberName(MsFormsModuleName, "QueryClose"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Resize = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Resize"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
            public static Declaration Terminate = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Terminate"), "MSForms.UserForm", null, true, false, Accessibility.Public, DeclarationType.Event);
        }
        #endregion
    }
}