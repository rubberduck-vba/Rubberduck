using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    internal static class AdodbObjectModel
    {
        private static IEnumerable<Declaration> _adodbDeclarations;
        private static readonly QualifiedModuleName AdodbModuleName = new QualifiedModuleName("ADODB", "ADODB");

        public static IEnumerable<Declaration> Declarations
        {
            get
            {
                if (_adodbDeclarations == null)
                {
                    var nestedTypes = typeof(AdodbObjectModel).GetNestedTypes(BindingFlags.NonPublic);
                    var fields = nestedTypes.SelectMany(t => t.GetFields());
                    var values = fields.Select(f => f.GetValue(null));
                    _adodbDeclarations = values.Cast<Declaration>();
                }

                return _adodbDeclarations;
            }
        }

        private class AdodbLib
        {
            public static readonly Declaration Adodb = new Declaration(new QualifiedMemberName(AdodbModuleName, "ADODB"), null, "ADODB", "ADODB", true, false, Accessibility.Global, DeclarationType.Project);

            public static readonly Declaration ADCPROP_ASYNCTHREADPRIORITY_ENUM = new Declaration(new QualifiedMemberName(AdodbModuleName, "ADCPROP_ASYNCTHREADPRIORITY_ENUM"), Adodb, "ADODB", "ADODB.ADCPROP_ASYNCTHREADPRIORITY_ENUM", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adPriorityAboveNormal = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adPriorityAboveNormal"), ADCPROP_ASYNCTHREADPRIORITY_ENUM, "ADODB", "ADODB.ADCPROP_ASYNCTHREADPRIORITY_ENUM", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration adPriorityBelowNormal = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adPriorityBelowNormal"), ADCPROP_ASYNCTHREADPRIORITY_ENUM, "ADODB", "ADODB.ADCPROP_ASYNCTHREADPRIORITY_ENUM", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration adPriorityHighest = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adPriorityHighest"), ADCPROP_ASYNCTHREADPRIORITY_ENUM, "ADODB", "ADODB.ADCPROP_ASYNCTHREADPRIORITY_ENUM", Accessibility.Global, DeclarationType.EnumerationMember, "5");
            public static Declaration adPriorityLowest = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adPriorityLowest"), ADCPROP_ASYNCTHREADPRIORITY_ENUM, "ADODB", "ADODB.ADCPROP_ASYNCTHREADPRIORITY_ENUM", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration adPriorityNormal = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adPriorityNormal"), ADCPROP_ASYNCTHREADPRIORITY_ENUM, "ADODB", "ADODB.ADCPROP_ASYNCTHREADPRIORITY_ENUM", Accessibility.Global, DeclarationType.EnumerationMember, "3");

            public static readonly Declaration ADCPROP_AUTORECALC_ENUM = new Declaration(new QualifiedMemberName(AdodbModuleName, "ADCPROP_AUTORECALC_ENUM"), Adodb, "ADODB", "ADODB.ADCPROP_AUTORECALC_ENUM", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adRecalcAlways = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adRecalcAlways"), ADCPROP_AUTORECALC_ENUM, "ADODB", "ADODB.ADCPROP_AUTORECALC_ENUM", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration adRecalcUpFront = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adRecalcUpFront"), ADCPROP_AUTORECALC_ENUM, "ADODB", "ADODB.ADCPROP_AUTORECALC_ENUM", Accessibility.Global, DeclarationType.EnumerationMember, "0");

            public static readonly Declaration ADCPROP_UPDATECRITERIA_ENUM = new Declaration(new QualifiedMemberName(AdodbModuleName, "ADCPROP_UPDATECRITERIA_ENUM"), Adodb, "ADODB", "ADODB.ADCPROP_UPDATECRITERIA_ENUM", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adCriteriaAllCols = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCriteriaAllCols"), ADCPROP_UPDATECRITERIA_ENUM, "ADODB", "ADODB.ADCPROP_UPDATECRITERIA_ENUM", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration adCriteriaKey = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCriteriaKey"), ADCPROP_UPDATECRITERIA_ENUM, "ADODB", "ADODB.ADCPROP_UPDATECRITERIA_ENUM", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration adCriteriaTimeStamp = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCriteriaTimeStamp"), ADCPROP_UPDATECRITERIA_ENUM, "ADODB", "ADODB.ADCPROP_UPDATECRITERIA_ENUM", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration adCriteriaUpdCols = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCriteriaUpdCols"), ADCPROP_UPDATECRITERIA_ENUM, "ADODB", "ADODB.ADCPROP_UPDATECRITERIA_ENUM", Accessibility.Global, DeclarationType.EnumerationMember, "2");

            public static readonly Declaration ADCPROP_UPDATERESYNC_ENUM = new Declaration(new QualifiedMemberName(AdodbModuleName, "ADCPROP_UPDATERESYNC_ENUM"), Adodb, "ADODB", "ADODB.ADCPROP_UPDATERESYNC_ENUM", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adResyncAll = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adResyncAll"), ADCPROP_UPDATERESYNC_ENUM, "ADODB", "ADODB.ADCPROP_UPDATERESYNC_ENUM", Accessibility.Global, DeclarationType.EnumerationMember, "15");
            public static Declaration adResyncAutoIncrement = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adResyncAutoIncrement"), ADCPROP_UPDATERESYNC_ENUM, "ADODB", "ADODB.ADCPROP_UPDATERESYNC_ENUM", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration adResyncConflicts = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adResyncConflicts"), ADCPROP_UPDATERESYNC_ENUM, "ADODB", "ADODB.ADCPROP_UPDATERESYNC_ENUM", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration adResyncInserts = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adResyncInserts"), ADCPROP_UPDATERESYNC_ENUM, "ADODB", "ADODB.ADCPROP_UPDATERESYNC_ENUM", Accessibility.Global, DeclarationType.EnumerationMember, "8");
            public static Declaration adResyncNone = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adResyncNone"), ADCPROP_UPDATERESYNC_ENUM, "ADODB", "ADODB.ADCPROP_UPDATERESYNC_ENUM", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration adResyncUpdates = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adResyncUpdates"), ADCPROP_UPDATERESYNC_ENUM, "ADODB", "ADODB.ADCPROP_UPDATERESYNC_ENUM", Accessibility.Global, DeclarationType.EnumerationMember, "4");

            public static readonly Declaration AffectEnum = new Declaration(new QualifiedMemberName(AdodbModuleName, "AffectEnum"), Adodb, "ADODB", "ADODB.AffectEnum", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adAffectAllChapters = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adAffectAllChapters"), AffectEnum, "ADODB", "ADODB.AffectEnum", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration adAffectCurrent = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adAffectCurrent"), AffectEnum, "ADODB", "ADODB.AffectEnum", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration adAffectGroup = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adAffectGroup"), AffectEnum, "ADODB", "ADODB.AffectEnum", Accessibility.Global, DeclarationType.EnumerationMember, "2");

            public static readonly Declaration BookmarkEnum = new Declaration(new QualifiedMemberName(AdodbModuleName, "BookmarkEnum"), Adodb, "ADODB", "ADODB.BookmarkEnum", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adBookmarkCurrent = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adBookmarkCurrent"), BookmarkEnum, "ADODB", "ADODB.BookmarkEnum", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration adBookmarkFirst = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adBookmarkFirst"), BookmarkEnum, "ADODB", "ADODB.BookmarkEnum", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration adBookmarkLast = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adBookmarkLast"), BookmarkEnum, "ADODB", "ADODB.BookmarkEnum", Accessibility.Global, DeclarationType.EnumerationMember, "2");

            public static readonly Declaration CommandTypeEnum = new Declaration(new QualifiedMemberName(AdodbModuleName, "CommandTypeEnum"), Adodb, "ADODB", "ADODB.CommandTypeEnum", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adCmdFile = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCmdFile"), CommandTypeEnum, "ADODB", "ADODB.CommandTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "256");
            public static Declaration adCmdStoredProc = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCmdStoredProc"), CommandTypeEnum, "ADODB", "ADODB.CommandTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration adCmdTable = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCmdTable"), CommandTypeEnum, "ADODB", "ADODB.CommandTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration adCmdTableDirect = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCmdTableDirect"), CommandTypeEnum, "ADODB", "ADODB.CommandTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "512");
            public static Declaration adCmdText = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCmdText"), CommandTypeEnum, "ADODB", "ADODB.CommandTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration adCmdUnknown = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCmdUnknown"), CommandTypeEnum, "ADODB", "ADODB.CommandTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "8");

            public static readonly Declaration CompareEnum = new Declaration(new QualifiedMemberName(AdodbModuleName, "CompareEnum"), Adodb, "ADODB", "ADODB.CompareEnum", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adCompareEqual = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCompareEqual"), CompareEnum, "ADODB", "ADODB.CompareEnum", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration adCompareGreaterThan = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCompareGreaterThan"), CompareEnum, "ADODB", "ADODB.CompareEnum", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration adCompareLessThan = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCompareLessThan"), CompareEnum, "ADODB", "ADODB.CompareEnum", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration adCompareNotComparable = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCompareNotComparable"), CompareEnum, "ADODB", "ADODB.CompareEnum", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration adCompareNotEqual = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCompareNotEqual"), CompareEnum, "ADODB", "ADODB.CompareEnum", Accessibility.Global, DeclarationType.EnumerationMember, "3");

            public static readonly Declaration ConnectModeEnum = new Declaration(new QualifiedMemberName(AdodbModuleName, "ConnectModeEnum"), Adodb, "ADODB", "ADODB.ConnectModeEnum", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adModeRead = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adModeRead"), ConnectModeEnum, "ADODB", "ADODB.ConnectModeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration adModeReadWrite = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adModeReadWrite"), ConnectModeEnum, "ADODB", "ADODB.ConnectModeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration adModeRecursive = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adModeRecursive"), ConnectModeEnum, "ADODB", "ADODB.ConnectModeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "4194304");
            public static Declaration adModeShareDenyNone = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adModeShareDenyNone"), ConnectModeEnum, "ADODB", "ADODB.ConnectModeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "16");
            public static Declaration adModeShareDenyRead = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adModeShareDenyRead"), ConnectModeEnum, "ADODB", "ADODB.ConnectModeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration adModeShareDenyWrite = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adModeShareDenyWrite"), ConnectModeEnum, "ADODB", "ADODB.ConnectModeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "8");
            public static Declaration adModeShareExclusive = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adModeShareExclusive"), ConnectModeEnum, "ADODB", "ADODB.ConnectModeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "12");
            public static Declaration adModeUnknown = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adModeUnknown"), ConnectModeEnum, "ADODB", "ADODB.ConnectModeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration adModeWrite = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adModeWrite"), ConnectModeEnum, "ADODB", "ADODB.ConnectModeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "2");

            public static readonly Declaration ConnectOptionEnum = new Declaration(new QualifiedMemberName(AdodbModuleName, "ConnectOptionEnum"), Adodb, "ADODB", "ADODB.ConnectOptionEnum", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adAsyncConnect = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adAsyncConnect"), ConnectModeEnum, "ADODB", "ADODB.ConnectOptionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "16");

            public static readonly Declaration ConnectPromptEnum = new Declaration(new QualifiedMemberName(AdodbModuleName, "ConnectPromptEnum"), Adodb, "ADODB", "ADODB.ConnectPromptEnum", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adPromptAlways = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adPromptAlways"), ConnectPromptEnum, "ADODB", "ADODB.ConnectPromptEnum", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration adPromptComplete = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adPromptComplete"), ConnectPromptEnum, "ADODB", "ADODB.ConnectPromptEnum", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration adPromptCompleteRequired = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adPromptCompleteRequired"), ConnectPromptEnum, "ADODB", "ADODB.ConnectPromptEnum", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration adPromptNever = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adPromptNever"), ConnectPromptEnum, "ADODB", "ADODB.ConnectPromptEnum", Accessibility.Global, DeclarationType.EnumerationMember, "4");

            public static readonly Declaration CopyRecordOptionsEnum = new Declaration(new QualifiedMemberName(AdodbModuleName, "CopyRecordOptionsEnum"), Adodb, "ADODB", "ADODB.CopyRecordOptionsEnum", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adCopyAllowEmulation = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCopyAllowEmulation"), CopyRecordOptionsEnum, "ADODB", "ADODB.CopyRecordOptionsEnum", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration adCopyNonRecursive = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCopyNonRecursive"), CopyRecordOptionsEnum, "ADODB", "ADODB.CopyRecordOptionsEnum", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration adCopyOverWrite = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCopyOverWrite"), CopyRecordOptionsEnum, "ADODB", "ADODB.CopyRecordOptionsEnum", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration adCopyUnspecified = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCopyUnspecified"), CopyRecordOptionsEnum, "ADODB", "ADODB.CopyRecordOptionsEnum", Accessibility.Global, DeclarationType.EnumerationMember, "-1");

            public static readonly Declaration CursorLocationEnum = new Declaration(new QualifiedMemberName(AdodbModuleName, "CursorLocationEnum"), Adodb, "ADODB", "ADODB.CursorLocationEnum", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adUseClient = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adUseClient"), CursorLocationEnum, "ADODB", "ADODB.CursorLocationEnum", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration adUseServer = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adUseServer"), CursorLocationEnum, "ADODB", "ADODB.CursorLocationEnum", Accessibility.Global, DeclarationType.EnumerationMember, "2");

            public static readonly Declaration CursorOptionEnum = new Declaration(new QualifiedMemberName(AdodbModuleName, "CursorOptionEnum"), Adodb, "ADODB", "ADODB.CursorOptionEnum", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adAddNew = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adAddNew"), CursorOptionEnum, "ADODB", "ADODB.CursorOptionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "16778240");
            public static Declaration adApproxPosition = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adApproxPosition"), CursorOptionEnum, "ADODB", "CursorOptionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "16384");
            public static Declaration adBookmark = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adBookmark"), CursorOptionEnum, "ADODB", "ADODB.CursorOptionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "8192");
            public static Declaration adDelete = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adDelete"), CursorOptionEnum, "ADODB", "ADODB.CursorOptionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "16779264");
            public static Declaration adFind = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adFind"), CursorOptionEnum, "ADODB", "ADODB.CursorOptionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "524288");
            public static Declaration adHoldRecords = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adHoldRecords"), CursorOptionEnum, "ADODB", "ADODB.CursorOptionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "256");
            public static Declaration adIndex = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adIndex"), CursorOptionEnum, "ADODB", "ADODB.CursorOptionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "8388608");
            public static Declaration adMovePrevious = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adMovePrevious"), CursorOptionEnum, "ADODB", "ADODB.CursorOptionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "512");
            public static Declaration adNotify = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adNotify"), CursorOptionEnum, "ADODB", "ADODB.CursorOptionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "262144");
            public static Declaration adResync = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adResync"), CursorOptionEnum, "ADODB", "ADODB.CursorOptionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "131072");
            public static Declaration adSeek = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adSeek"), CursorOptionEnum, "ADODB", "ADODB.CursorOptionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "4194304");
            public static Declaration adUpdate = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adUpdate"), CursorOptionEnum, "ADODB", "ADODB.CursorOptionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "16809984");
            public static Declaration adUpdateBatch = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adUpdateBatch"), CursorOptionEnum, "ADODB", "ADODB.CursorOptionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "65536");

            public static readonly Declaration CursorTypeEnum = new Declaration(new QualifiedMemberName(AdodbModuleName, "CursorTypeEnum"), Adodb, "ADODB", "ADODB.CursorTypeEnum", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adOpenDynamic = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adOpenDynamic"), CursorTypeEnum, "ADODB", "ADODB.CursorTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration adOpenForwardOnly = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adOpenForwardOnly"), CursorTypeEnum, "ADODB", "ADODB.CursorTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration adOpenKeyset = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adOpenKeyset"), CursorTypeEnum, "ADODB", "ADODB.CursorTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration adOpenStatic = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adOpenStatic"), CursorTypeEnum, "ADODB", "ADODB.CursorTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "3");

            public static readonly Declaration DataTypeEnum = new Declaration(new QualifiedMemberName(AdodbModuleName, "DataTypeEnum"), Adodb, "ADODB", "ADODB.DataTypeEnum", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adArray = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adArray"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "8192");
            public static Declaration adBigInt = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adBigInt"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "20");
            public static Declaration adBinary = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adBinary"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "128");
            public static Declaration adBoolean = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adBoolean"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "11");
            public static Declaration adBSTR = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adBSTR"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "8");
            public static Declaration adChapter = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adChapter"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "136");
            public static Declaration adChar = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adChar"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "129");
            public static Declaration adCurrency = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adCurrency"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "6");
            public static Declaration adDate = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adDate"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "7");
            public static Declaration adDBDate = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adDBDate"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "133");
            public static Declaration adDBTime = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adDBTime"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "134");
            public static Declaration adDBTimeStamp = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adDBTimeStamp"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "135");
            public static Declaration adDecimal = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adDecimal"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "14");
            public static Declaration adEmpty = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adEmpty"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "0");
            public static Declaration adError = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adError"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "10");
            public static Declaration adFileTime = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adFileTime"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "64");
            public static Declaration adGUID = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adGUID"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "72");
            public static Declaration adIDispatch = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adIDispatch"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "9");
            public static Declaration adInteger = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adInteger"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration adIUnknown = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adIUnknown"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "13");
            public static Declaration adLongVarBinary = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adLongVarBinary"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "205");
            public static Declaration adLongVarChar = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adLongVarChar"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "201");
            public static Declaration adLongVarWChar = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adLongVarWChar"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "203");
            public static Declaration adNumeric = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adNumeric"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "131");
            public static Declaration adPropVariant = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adPropVariant"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "138");
            public static Declaration adSingle = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adSingle"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration adSmallInt = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adSmallInt"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration adTinyInt = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adTinyInt"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "16");
            public static Declaration adUnsignedBigInt = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adUnsignedBigInt"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "21");
            public static Declaration adUnsignedInt = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adUnsignedInt"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "19");
            public static Declaration adUnsignedSmallInt = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adUnsignedSmallInt"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "18");
            public static Declaration adUnsignedTinyInt = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adUnsignedTinyInt"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "17");
            public static Declaration adUserDefined = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adUserDefined"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "132");
            public static Declaration adVarBinary = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adVarBinary"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "204");
            public static Declaration adVarChar = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adVarChar"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "200");
            public static Declaration adVariant = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adVariant"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "12");
            public static Declaration adVarNumeric = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adVarNumeric"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "139");
            public static Declaration adVarWChar = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adVarWChar"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "202");
            public static Declaration adWChar = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adWChar"), DataTypeEnum, "ADODB", "ADODB.DataTypeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "130");

            public static readonly Declaration EditModeEnum = new Declaration(new QualifiedMemberName(AdodbModuleName, "EditModeEnum"), Adodb, "ADODB", "ADODB.EditModeEnum", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adEditAdd = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adArray"), EditModeEnum, "ADODB", "ADODB.EditModeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration adEditDelete = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adBigInt"), EditModeEnum, "ADODB", "ADODB.EditModeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration adEditInProgress = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adBinary"), EditModeEnum, "ADODB", "ADODB.EditModeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration adEditNone = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adBoolean"), EditModeEnum, "ADODB", "ADODB.EditModeEnum", Accessibility.Global, DeclarationType.EnumerationMember, "0");

            /*
             * TODO:
             * ErrorValueEnum
             * EventReasonEnum
             * EventStatusEnum
             * ExecuteOptionEnum
             * Field
             * FieldAttributeEnum
             * FieldEnum
             * Fields
             * FieldStatusEnum
             * FilterGroupEnum
             * GetRowsOptionEnum
             * IsolationLevelEnum
             * LineSeparatorEnum
             * LockTypeEnum
             * MarshalOptionsEnum
             * MoveRecordOptionsEnum
             * ObjectStateEnum
             * 
             */

            public static readonly Declaration ParameterAttributesEnum = new Declaration(new QualifiedMemberName(AdodbModuleName, "ParameterAttributesEnum"), Adodb, "ADODB", "ADODB.ParameterAttributesEnum", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adParamLong = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adParamLong"), ParameterAttributesEnum, "ADODB", "ADODB.ParameterAttributesEnum", Accessibility.Global, DeclarationType.EnumerationMember, "128");
            public static Declaration adParamNullable = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adParamNullable"), ParameterAttributesEnum, "ADODB", "ADODB.ParameterAttributesEnum", Accessibility.Global, DeclarationType.EnumerationMember, "64");
            public static Declaration adParamSigned = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adParamSigned"), ParameterAttributesEnum, "ADODB", "ADODB.ParameterAttributesEnum", Accessibility.Global, DeclarationType.EnumerationMember, "16");

            public static readonly Declaration ParameterDirectionEnum = new Declaration(new QualifiedMemberName(AdodbModuleName, "ParameterDirectionEnum"), Adodb, "ADODB", "ADODB.ParameterDirectionEnum", false, false, Accessibility.Global, DeclarationType.Enumeration);
            public static Declaration adParamInput = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adParamInput"), ParameterDirectionEnum, "ADODB", "ADODB.ParameterDirectionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "1");
            public static Declaration adParamInputOutput = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adParamInputOutput"), ParameterDirectionEnum, "ADODB", "ADODB.ParameterDirectionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "3");
            public static Declaration adParamOutput = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adParamOutput"), ParameterDirectionEnum, "ADODB", "ADODB.ParameterDirectionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "2");
            public static Declaration adParamReturnValue = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adParamReturnValue"), ParameterDirectionEnum, "ADODB", "ADODB.ParameterDirectionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "4");
            public static Declaration adParamUnknown = new ValuedDeclaration(new QualifiedMemberName(AdodbModuleName, "adParamUnknown"), ParameterDirectionEnum, "ADODB", "ADODB.ParameterDirectionEnum", Accessibility.Global, DeclarationType.EnumerationMember, "0");
        }

        private class CommandClass
        {
            private static readonly QualifiedModuleName CommandModuleName = new QualifiedModuleName("ADODB", "Command");
            public static readonly Declaration Command = new Declaration(new QualifiedMemberName(CommandModuleName, "Command"), AdodbLib.Adodb, "ADODB", "ADODB.Command", true, false, Accessibility.Global, DeclarationType.Class);

            public static readonly Declaration ActiveConnectionGet = new Declaration(new QualifiedMemberName(CommandModuleName, "ActiveConnection"), Command, "ADODB.Command", "Connection", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration ActiveConnectionSet = new Declaration(new QualifiedMemberName(CommandModuleName, "ActiveConnection"), Command, "ADODB.Command", null, false, false, Accessibility.Public, DeclarationType.PropertySet);

            public static readonly Declaration Cancel = new Declaration(new QualifiedMemberName(CommandModuleName, "Cancel"), Command, "ADODB.Command", null, false, false, Accessibility.Public, DeclarationType.Procedure);

            public static readonly Declaration CommandStreamGet = new Declaration(new QualifiedMemberName(CommandModuleName, "CommandStream"), Command, "ADODB.Command", "Stream", false, false, Accessibility.Public, DeclarationType.PropertyGet); // cheating on return type (actually Variant)
            public static readonly Declaration CommandStreamSet = new Declaration(new QualifiedMemberName(CommandModuleName, "CommandStream"), Command, "ADODB.Command", null, false, false, Accessibility.Public, DeclarationType.PropertySet);

            public static readonly Declaration CommandTextGet = new Declaration(new QualifiedMemberName(CommandModuleName, "CommandText"), Command, "ADODB.Command", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration CommandTextLet = new Declaration(new QualifiedMemberName(CommandModuleName, "CommandText"), Command, "ADODB.Command", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration CommandTimeoutGet = new Declaration(new QualifiedMemberName(CommandModuleName, "CommandTimeout"), Command, "ADODB.Command", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration CommandTimeoutLet = new Declaration(new QualifiedMemberName(CommandModuleName, "CommandTimeout"), Command, "ADODB.Command", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration CommandTypeGet = new Declaration(new QualifiedMemberName(CommandModuleName, "CommandType"), Command, "ADODB.Command", "CommandTypeEnum", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration CommandTypeSet = new Declaration(new QualifiedMemberName(CommandModuleName, "CommandType"), Command, "ADODB.Command", null, false, false, Accessibility.Public, DeclarationType.PropertySet);

            public static readonly Declaration CreateParameter = new Declaration(new QualifiedMemberName(CommandModuleName, "CreateParameter"), Command, "ADODB.Command", "Parameter", false, false, Accessibility.Public, DeclarationType.Function);

            public static readonly Declaration CommandDialectGet = new Declaration(new QualifiedMemberName(CommandModuleName, "CommandDialect"), Command, "ADODB.Command", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration CommandDialectLet = new Declaration(new QualifiedMemberName(CommandModuleName, "CommandDialect"), Command, "ADODB.Command", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration Execute = new Declaration(new QualifiedMemberName(CommandModuleName, "Execute"), Command, "ADODB.Command", "Recordset", false, false, Accessibility.Public, DeclarationType.Function);

            public static readonly Declaration NameGet = new Declaration(new QualifiedMemberName(CommandModuleName, "Name"), Command, "ADODB.Command", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration NameLet = new Declaration(new QualifiedMemberName(CommandModuleName, "Name"), Command, "ADODB.Command", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration NamedParametersGet = new Declaration(new QualifiedMemberName(CommandModuleName, "NamedParameters"), Command, "ADODB.Command", "Boolean", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration NamedParametersLet = new Declaration(new QualifiedMemberName(CommandModuleName, "NamedParameters"), Command, "ADODB.Command", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration ParametersGet = new Declaration(new QualifiedMemberName(CommandModuleName, "Parameters"), Command, "ADODB.Command", "Parameters", false, false, Accessibility.Public, DeclarationType.PropertyGet);

            public static readonly Declaration PreparedGet = new Declaration(new QualifiedMemberName(CommandModuleName, "Prepared"), Command, "ADODB.Command", "Boolean", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration PreparedLet = new Declaration(new QualifiedMemberName(CommandModuleName, "Prepared"), Command, "ADODB.Command", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration PropertiesGet = new Declaration(new QualifiedMemberName(CommandModuleName, "Properties"), Command, "ADODB.Command", "Properties", false, false, Accessibility.Public, DeclarationType.PropertyGet);

            public static readonly Declaration StateGet = new Declaration(new QualifiedMemberName(CommandModuleName, "State"), Command, "ADODB.Command", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
        }

        private class ConnectionClass
        {
            private static readonly QualifiedModuleName ConnectionModuleName = new QualifiedModuleName("ADODB", "Connection");
            public static readonly Declaration Connection = new Declaration(new QualifiedMemberName(ConnectionModuleName, "Connection"), AdodbLib.Adodb, "ADODB", "ADODB.Connection", true, false, Accessibility.Global, DeclarationType.Class);

            public static readonly Declaration ConnectionTimeoutGet = new Declaration(new QualifiedMemberName(ConnectionModuleName, "ConnectionTimeout"), Connection, "ADODB.Connection", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration ConnectionTimeoutLet = new Declaration(new QualifiedMemberName(ConnectionModuleName, "ConnectionTimeout"), Connection, "ADODB.Connection", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration Close = new Declaration(new QualifiedMemberName(ConnectionModuleName, "Close"), Connection, "ADODB.Connection", null, false, false, Accessibility.Public, DeclarationType.Procedure);

            public static readonly Declaration CursorLocationGet = new Declaration(new QualifiedMemberName(ConnectionModuleName, "CursorLocation"), Connection, "ADODB.Connection", "CursorLocationEnum", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration CursorLocationLet = new Declaration(new QualifiedMemberName(ConnectionModuleName, "CursorLocation"), Connection, "ADODB.Connection", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration DefaultDatabaseGet = new Declaration(new QualifiedMemberName(ConnectionModuleName, "DefaultDatabase"), Connection, "ADODB.Connection", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration DefaultDatabaseLet = new Declaration(new QualifiedMemberName(ConnectionModuleName, "DefaultDatabase"), Connection, "ADODB.Connection", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration Disconnect = new Declaration(new QualifiedMemberName(ConnectionModuleName, "Disconnect"), Connection, "ADODB.Connection", null, false, false, Accessibility.Public, DeclarationType.Event);

            public static readonly Declaration Errors = new Declaration(new QualifiedMemberName(ConnectionModuleName, "Errors"), Connection, "ADODB.Connection", "Errors", false, false, Accessibility.Public, DeclarationType.PropertyGet);

            public static readonly Declaration Execute = new Declaration(new QualifiedMemberName(ConnectionModuleName, "Execute"), Connection, "ADODB.Connection", "Recordset", false, false, Accessibility.Public, DeclarationType.Function);
            public static readonly Declaration ExecuteComplete = new Declaration(new QualifiedMemberName(ConnectionModuleName, "ExecuteComplete"), Connection, "ADODB.Connection", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static readonly Declaration InfoMessage = new Declaration(new QualifiedMemberName(ConnectionModuleName, "InfoMessage"), Connection, "ADODB.Connection", null, false, false, Accessibility.Public, DeclarationType.Event);

            public static readonly Declaration IsolationLevelGet = new Declaration(new QualifiedMemberName(ConnectionModuleName, "IsolationLevel"), Connection, "ADODB.Connection", "IsolationLevelEnum", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration IsolationLevelLet = new Declaration(new QualifiedMemberName(ConnectionModuleName, "IsolationLevel"), Connection, "ADODB.Connection", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration ModeGet = new Declaration(new QualifiedMemberName(ConnectionModuleName, "Mode"), Connection, "ADODB.Connection", "ConnectModeEnum", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration ModeLet = new Declaration(new QualifiedMemberName(ConnectionModuleName, "Mode"), Connection, "ADODB.Connection", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration Open = new Declaration(new QualifiedMemberName(ConnectionModuleName, "Open"), Connection, "ADODB.Connection", null, false, false, Accessibility.Public, DeclarationType.Procedure);
            public static readonly Declaration OpenSchema = new Declaration(new QualifiedMemberName(ConnectionModuleName, "OpenSchema"), Connection, "ADODB.Connection", "Recordset", false, false, Accessibility.Public, DeclarationType.Function);

            public static readonly Declaration PropertiesGet = new Declaration(new QualifiedMemberName(ConnectionModuleName, "Properties"), Connection, "ADODB.Connection", "Properties", false, false, Accessibility.Public, DeclarationType.PropertyGet);

            public static readonly Declaration ProviderGet = new Declaration(new QualifiedMemberName(ConnectionModuleName, "Provider"), Connection, "ADODB.Connection", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration ProviderLet = new Declaration(new QualifiedMemberName(ConnectionModuleName, "Provider"), Connection, "ADODB.Connection", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration RollbackTrans = new Declaration(new QualifiedMemberName(ConnectionModuleName, "RollbackTrans"), Connection, "ADODB.Connection", null, false, false, Accessibility.Public, DeclarationType.Procedure);
            public static readonly Declaration RollbackTransComplete = new Declaration(new QualifiedMemberName(ConnectionModuleName, "RollbackTransComplete"), Connection, "ADODB.Connection", null, false, false, Accessibility.Public, DeclarationType.Event);

            public static readonly Declaration StateGet = new Declaration(new QualifiedMemberName(ConnectionModuleName, "State"), Connection, "ADODB.Connection", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration VersionGet = new Declaration(new QualifiedMemberName(ConnectionModuleName, "Version"), Connection, "ADODB.Connection", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet);

            public static readonly Declaration WillConnect = new Declaration(new QualifiedMemberName(ConnectionModuleName, "WillConnect"), Connection, "ADODB.Connection", null, false, false, Accessibility.Public, DeclarationType.Event);
            public static readonly Declaration WillExecute = new Declaration(new QualifiedMemberName(ConnectionModuleName, "WillExecute"), Connection, "ADODB.Connection", null, false, false, Accessibility.Public, DeclarationType.Event);
        }

        private class ErrorsClass
        {
            private static readonly QualifiedModuleName ErrorsModuleName = new QualifiedModuleName("ADODB", "Errors");
            public static readonly Declaration Errors = new Declaration(new QualifiedMemberName(ErrorsModuleName, "Errors"), AdodbLib.Adodb, "ADODB", "ADODB.Errors", true, false, Accessibility.Global, DeclarationType.Class);

            public static readonly Declaration Clear = new Declaration(new QualifiedMemberName(ErrorsModuleName, "Clear"), Errors, "ADODB.Errors", null, false, false, Accessibility.Public, DeclarationType.Procedure);

            public static readonly Declaration Count = new Declaration(new QualifiedMemberName(ErrorsModuleName, "Count"), Errors, "ADODB.Errors", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration Item = new Declaration(new QualifiedMemberName(ErrorsModuleName, "Item"), Errors, "ADODB.Errors", "Error", false, false, Accessibility.Public, DeclarationType.PropertyGet);

            public static readonly Declaration Refresh = new Declaration(new QualifiedMemberName(ErrorsModuleName, "Refresh"), Errors, "ADODB.Errors", null, false, false, Accessibility.Public, DeclarationType.Procedure);
        }

        private class ErrorClass
        {
            private static readonly QualifiedModuleName ErrorModuleName = new QualifiedModuleName("ADODB", "Error");
            public static readonly Declaration Error = new Declaration(new QualifiedMemberName(ErrorModuleName, "Error"), AdodbLib.Adodb, "ADODB", "ADODB.Error", true, false, Accessibility.Global, DeclarationType.Class);

            public static readonly Declaration Description = new Declaration(new QualifiedMemberName(ErrorModuleName, "Description"), Error, "ADODB.Error", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration HelpContext = new Declaration(new QualifiedMemberName(ErrorModuleName, "HelpContext"), Error, "ADODB.Error", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration HelpFile = new Declaration(new QualifiedMemberName(ErrorModuleName, "HelpFile"), Error, "ADODB.Error", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration NativeError = new Declaration(new QualifiedMemberName(ErrorModuleName, "NativeError"), Error, "ADODB.Error", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration Number = new Declaration(new QualifiedMemberName(ErrorModuleName, "Number"), Error, "ADODB.Error", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration Source = new Declaration(new QualifiedMemberName(ErrorModuleName, "Source"), Error, "ADODB.Error", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration SQLState = new Declaration(new QualifiedMemberName(ErrorModuleName, "SQLState"), Error, "ADODB.Error", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet);
        }

        /*
         * TODO:
         * Field (hidden type)
         * Fields (hidden type)
         */

        private class ParameterClass
        {
            private static readonly QualifiedModuleName ParameterModuleName = new QualifiedModuleName("ADODB", "Parameter");
            public static readonly Declaration Parameter = new Declaration(new QualifiedMemberName(ParameterModuleName, "Parameter"), AdodbLib.Adodb, "ADODB", "ADODB.Parameter", true, false, Accessibility.Global, DeclarationType.Class);

            public static readonly Declaration AppendChunk = new Declaration(new QualifiedMemberName(ParameterModuleName, "AppendChunk"), Parameter, "ADODB.Parameter", null, false, false, Accessibility.Public, DeclarationType.Procedure);

            public static readonly Declaration AttributesGet = new Declaration(new QualifiedMemberName(ParameterModuleName, "Attributes"), Parameter, "ADODB.Parameter", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration AttributesLet = new Declaration(new QualifiedMemberName(ParameterModuleName, "Attributes"), Parameter, "ADODB.Parameter", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);
            
            public static readonly Declaration DirectionGet = new Declaration(new QualifiedMemberName(ParameterModuleName, "Direction"), Parameter, "ADODB.Parameter", "ParameterDirectionEnum", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration DirectionLet = new Declaration(new QualifiedMemberName(ParameterModuleName, "Direction"), Parameter, "ADODB.Parameter", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration NameGet = new Declaration(new QualifiedMemberName(ParameterModuleName, "Name"), Parameter, "ADODB.Parameter", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration NameLet = new Declaration(new QualifiedMemberName(ParameterModuleName, "Name"), Parameter, "ADODB.Parameter", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration NumericScaleGet = new Declaration(new QualifiedMemberName(ParameterModuleName, "NumericScale"), Parameter, "ADODB.Parameter", "Byte", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration NumericScaleLet = new Declaration(new QualifiedMemberName(ParameterModuleName, "NumericScale"), Parameter, "ADODB.Parameter", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration PrecisionGet = new Declaration(new QualifiedMemberName(ParameterModuleName, "Precision"), Parameter, "ADODB.Parameter", "Byte", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration PrecisionLet = new Declaration(new QualifiedMemberName(ParameterModuleName, "Precision"), Parameter, "ADODB.Parameter", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration PropertiesGet = new Declaration(new QualifiedMemberName(ParameterModuleName, "Properties"), Parameter, "ADODB.Parameter", "Properties", false, false, Accessibility.Public, DeclarationType.PropertyGet);

            public static readonly Declaration SizeGet = new Declaration(new QualifiedMemberName(ParameterModuleName, "Size"), Parameter, "ADODB.Parameter", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration SizeLet = new Declaration(new QualifiedMemberName(ParameterModuleName, "Size"), Parameter, "ADODB.Parameter", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration TypeGet = new Declaration(new QualifiedMemberName(ParameterModuleName, "Type"), Parameter, "ADODB.Parameter", "DataTypeEnum", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration TypeLet = new Declaration(new QualifiedMemberName(ParameterModuleName, "Type"), Parameter, "ADODB.Parameter", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration ValueGet = new Declaration(new QualifiedMemberName(ParameterModuleName, "Value"), Parameter, "ADODB.Parameter", "Variant", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration ValueLet = new Declaration(new QualifiedMemberName(ParameterModuleName, "Value"), Parameter, "ADODB.Parameter", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);
        }

        private class ParametersClass
        {
            private static readonly QualifiedModuleName ParametersModuleName = new QualifiedModuleName("ADODB", "Parameters");
            public static readonly Declaration Parameters = new Declaration(new QualifiedMemberName(ParametersModuleName, "Parameters"), AdodbLib.Adodb, "ADODB", "ADODB.Errors", true, false, Accessibility.Global, DeclarationType.Class);

            public static readonly Declaration Append = new Declaration(new QualifiedMemberName(ParametersModuleName, "Append"), Parameters, "ADODB.Parameters", null, false, false, Accessibility.Public, DeclarationType.Procedure);
            public static readonly Declaration Clear = new Declaration(new QualifiedMemberName(ParametersModuleName, "Clear"), Parameters, "ADODB.Parameters", null, false, false, Accessibility.Public, DeclarationType.Procedure);

            public static readonly Declaration Count = new Declaration(new QualifiedMemberName(ParametersModuleName, "Count"), Parameters, "ADODB.Parameters", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration Item = new Declaration(new QualifiedMemberName(ParametersModuleName, "Item"), Parameters, "ADODB.Parameters", "Parameter", false, false, Accessibility.Public, DeclarationType.PropertyGet);

            public static readonly Declaration Refresh = new Declaration(new QualifiedMemberName(ParametersModuleName, "Refresh"), Parameters, "ADODB.Parameters", null, false, false, Accessibility.Public, DeclarationType.Procedure);
        }

        private class PropertiesClass
        {
            private static readonly QualifiedModuleName PropertiesModuleName = new QualifiedModuleName("ADODB", "Properties");
            public static readonly Declaration Properties = new Declaration(new QualifiedMemberName(PropertiesModuleName, "Properties"), AdodbLib.Adodb, "ADODB", "ADODB.Properties", true, false, Accessibility.Global, DeclarationType.Class);

            public static readonly Declaration Count = new Declaration(new QualifiedMemberName(PropertiesModuleName, "Count"), Properties, "ADODB.Properties", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration Item = new Declaration(new QualifiedMemberName(PropertiesModuleName, "Item"), Properties, "ADODB.Properties", "Parameter", false, false, Accessibility.Public, DeclarationType.PropertyGet);

            public static readonly Declaration Refresh = new Declaration(new QualifiedMemberName(PropertiesModuleName, "Refresh"), Properties, "ADODB.Properties", null, false, false, Accessibility.Public, DeclarationType.Procedure);
        }

        private class PropertyClass
        {
            private static readonly QualifiedModuleName PropertyModuleName = new QualifiedModuleName("ADODB", "Property");
            public static readonly Declaration Property = new Declaration(new QualifiedMemberName(PropertyModuleName, "Property"), AdodbLib.Adodb, "ADODB", "ADODB.Property", true, false, Accessibility.Global, DeclarationType.Class);

            public static readonly Declaration AttributesGet = new Declaration(new QualifiedMemberName(PropertyModuleName, "Attributes"), Property, "ADODB.Property", "Long", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration AttributesLet = new Declaration(new QualifiedMemberName(PropertyModuleName, "Attributes"), Property, "ADODB.Property", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration NameGet = new Declaration(new QualifiedMemberName(PropertyModuleName, "Name"), Property, "ADODB.Property", "String", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration NameLet = new Declaration(new QualifiedMemberName(PropertyModuleName, "Name"), Property, "ADODB.Property", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration TypeGet = new Declaration(new QualifiedMemberName(PropertyModuleName, "Type"), Property, "ADODB.Property", "DataTypeEnum", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration TypeLet = new Declaration(new QualifiedMemberName(PropertyModuleName, "Type"), Property, "ADODB.Property", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);

            public static readonly Declaration ValueGet = new Declaration(new QualifiedMemberName(PropertyModuleName, "Value"), Property, "ADODB.Property", "Variant", false, false, Accessibility.Public, DeclarationType.PropertyGet);
            public static readonly Declaration ValueLet = new Declaration(new QualifiedMemberName(PropertyModuleName, "Value"), Property, "ADODB.Property", null, false, false, Accessibility.Public, DeclarationType.PropertyLet);
        }
    }
}
