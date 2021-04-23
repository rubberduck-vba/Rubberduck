using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;

namespace Rubberduck.Resources
{
    /// <summary>
    /// Identifies a static <see cref="Tokens"/> string that isn't a legal identifier name for user code, e.g. keyword or reserved identifier.
    /// </summary>
    [AttributeUsage(AttributeTargets.Field)]
    public class ForbiddenAttribute : Attribute { }

    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public static class Tokens
    {
        public static IEnumerable<string> IllegalIdentifierNames =>
            typeof(Tokens).GetFields().Where(item => item.GetCustomAttributes<ForbiddenAttribute>().Any())
                .Select(item => item.GetValue(null)).Cast<string>().Select(item => item);

        [Forbidden]
        public static readonly string Abs = "Abs";
        [Forbidden]
        public static readonly string AddressOf = "AddressOf";
        [Forbidden]
        public static readonly string And = "And";
        [Forbidden]
        public static readonly string Any = "Any";
        [Forbidden]
        public static readonly string As = "As";
        public static readonly string Asc = "Asc";
        [Forbidden]
        public static readonly string Attribute = "Attribute";
        [Forbidden]
        public static readonly string Array = "Array";
        public static readonly string Base = "Base";
        public static readonly string Beep = "Beep";
        public static readonly string Binary = "Binary";
        [Forbidden]
        public static readonly string Boolean = "Boolean";
        [Forbidden]
        public static readonly string ByRef = "ByRef";
        [Forbidden]
        public static readonly string Byte = "Byte";
        [Forbidden]
        public static readonly string ByVal = "ByVal";
        [Forbidden]
        public static readonly string Call = "Call";
        [Forbidden]
        public static readonly string Case = "Case";
        [Forbidden]
        public static readonly string CBool = "CBool";
        [Forbidden]
        public static readonly string CByte = "CByte";
        [Forbidden]
        public static readonly string CCur = "CCur";
        [Forbidden]
        public static readonly string CDate = "CDate";
        [Forbidden]
        public static readonly string CDbl = "CDbl";
        [Forbidden]
        public static readonly string CDec = "CDec";
        public static readonly string ChDir = "ChDir";
        public static readonly string ChDrive = "ChDrive";
        public static readonly string Chr = "Chr";
        public static readonly string ChrB = "ChrB";
        public static readonly string ChrW = "ChrW";
        [Forbidden]
        public static readonly string CInt = "CInt";
        [Forbidden]
        public static readonly string CLng = "CLng";
        public static readonly string CLngLng = "CLngLng";
        [Forbidden]
        public static readonly string CLngPtr = "CLngPtr";
        public static readonly string Close = "Close";
        public static readonly string Command = "Command";
        [Forbidden]
        public static readonly string CommentMarker = "'";
        public static readonly string Compare = "Compare";
        [Forbidden]
        public static readonly string Const = "Const";
        public static readonly string Cos = "Cos";
        [Forbidden]
        public static readonly string CSng = "CSng";
        [Forbidden]
        public static readonly string CStr = "CStr";
        public static readonly string CurDir = "CurDir";
        [Forbidden]
        public static readonly string Currency = "Currency";
        [Forbidden]
        public static readonly string CVar = "CVar";
        [Forbidden]
        public static readonly string CVErr = "CVErr";
        public static readonly string Data = "Data";
        [Forbidden]
        public static readonly string Date = "Date";
        public static readonly string DateValue = "DateValue";
        public static readonly string Day = "Day";
        [Forbidden]
        public static readonly string Debug = "Debug";
        [Forbidden]
        public static readonly string Decimal = "Decimal";
        [Forbidden]
        public static readonly string Declare = "Declare";
        [Forbidden]
        public static readonly string DefBool = "DefBool";
        [Forbidden]
        public static readonly string DefByte = "DefByte";
        [Forbidden]
        public static readonly string DefCur = "DefCur";
        [Forbidden]
        public static readonly string DefDate = "DefDate";
        [Forbidden]
        public static readonly string DefDbl = "DefDbl";
        [Forbidden]
        public static readonly string DefInt = "DefInt";
        [Forbidden]
        public static readonly string DefLng = "DefLng";
        public static readonly string DefLngLng = "DefLngLng";
        [Forbidden]
        public static readonly string DefLngPtr = "DefLngptr";
        [Forbidden]
        public static readonly string DefObj = "DefObj";
        [Forbidden]
        public static readonly string DefSng = "DefSng";
        [Forbidden]
        public static readonly string DefStr = "DefStr";
        [Forbidden]
        public static readonly string DefVar = "DefVar";
        [Forbidden]
        public static readonly string Dim = "Dim";
        public static readonly string Dir = "Dir";
        [Forbidden]
        public static readonly string Do = "Do";
        [Forbidden]
        public static readonly string DoEvents = "DoEvents";
        [Forbidden]
        public static readonly string Double = "Double";
        [Forbidden]
        public static readonly string Each = "Each";
        [Forbidden]
        public static readonly string Else = "Else";
        [Forbidden]
        public static readonly string ElseIf = "ElseIf";
        [Forbidden]
        public static readonly string Empty = "Empty";
        [Forbidden]
        public static readonly string End = "End";
        [Forbidden]
        public static readonly string Enum = "Enum";
        public static readonly string Environ = "Environ";
        public static readonly string EOF = "EOF";
        [Forbidden]
        public static readonly string Eqv = "Eqv";
        [Forbidden]
        public static readonly string Erase = "Erase";
        public static readonly string Err = "Err";
        public static readonly string Error = "Error";
        [Forbidden]
        public static readonly string Event = "Event";
        [Forbidden]
        public static readonly string Exit = "Exit";
        public static readonly string Exp = "Exp";
        public static readonly string Explicit = "Explicit";
        [Forbidden]
        public static readonly string False = "False";
        [Forbidden]
        public static readonly string Fix = "Fix";
        [Forbidden]
        public static readonly string For = "For";
        public static readonly string Format = "Format";
        public static readonly string FreeFile = "FreeFile";
        [Forbidden]
        public static readonly string Friend = "Friend";
        [Forbidden]
        public static readonly string Function = "Function";
        [Forbidden]
        public static readonly string Get = "Get";
        [Forbidden]
        public static readonly string Global = "Global";
        [Forbidden]
        public static readonly string GoSub = "GoSub";
        [Forbidden]
        public static readonly string GoTo = "GoTo";
        public static readonly string Hex = "Hex";
        public static readonly string Hour = "Hour";
        public static readonly string IDispatch = "IDispatch";
        [Forbidden]
        public static readonly string If = "If";
        [Forbidden]
        public static readonly string Imp = "Imp";
        [Forbidden]
        public static readonly string Implements = "Implements";
        [Forbidden]
        public static readonly string In = "In";
        [Forbidden]
        public static readonly string Input = "Input";
        [Forbidden]
        public static readonly string InputB = "InputB";
        public static readonly string InputBox = "InputBox";
        public static readonly string InStr = "InStr";
        [Forbidden]
        public static readonly string Int = "Int";
        [Forbidden]
        public static readonly string Integer = "Integer";
        [Forbidden]
        public static readonly string Is = "Is";
        public static readonly string IsDate = "IsDate";
        public static readonly string IsEmpty = "IsEmpty";
        public static readonly string IsNull = "IsNull";
        public static readonly string IsNumeric = "IsNumeric";
        public static readonly string Join = "Join";
        public static readonly string Kill = "Kill";
        [Forbidden]
        public static readonly string LBound = "LBound";
        public static readonly string LCase = "LCase";
        public static readonly string Left = "Left";
        public static readonly string LeftB = "LeftB";
        [Forbidden]
        public static readonly string Len = "Len";
        [Forbidden]
        public static readonly string LenB = "LenB";
        [Forbidden]
        public static readonly string Let = "Let";
        [Forbidden]
        public static readonly string Like = "Like";
        public static readonly string Line = "Line";
        [Forbidden]
        public static readonly string LineContinuation = " _";
        [Forbidden]
        public static readonly string Lock = "Lock";
        public static readonly string LOF = "LOF";
        [Forbidden]
        public static readonly string Long = "Long";
        public static readonly string LongLong = "LongLong";
        [Forbidden]
        public static readonly string LongPtr = "LongPtr";
        [Forbidden]
        public static readonly string Loop = "Loop";
        [Forbidden]
        public static readonly string LSet = "LSet";
        public static readonly string LTrim = "LTrim";
        [Forbidden]
        public static readonly string Me = "Me";
        public static readonly string Mid = "Mid";
        public static readonly string MidB = "MidB";
        public static readonly string Minute = "Minute";
        public static readonly string MkDir = "MkDir";
        [Forbidden]
        public static readonly string Mod = "Mod";
        public static readonly string Month = "Month";
        public static readonly string MsgBox = "MsgBox";
        [Forbidden]
        public static readonly string New = "New";
        [Forbidden]
        public static readonly string Next = "Next";
        [Forbidden]
        public static readonly string Not = "Not";
        [Forbidden]
        public static readonly string Nothing = "Nothing";
        public static readonly string Now = "Now";
        [Forbidden]
        public static readonly string Null = "Null";
        public static readonly string Object = "Object";
        public static readonly string Oct = "Oct";
        [Forbidden]
        public static readonly string On = "On";
        [Forbidden]
        public static readonly string Open = "Open";
        [Forbidden]
        public static readonly string Option = "Option";
        [Forbidden]
        public static readonly string Optional = "Optional";
        [Forbidden]
        public static readonly string Or = "Or";
        public static readonly string Output = "Output";
        [Forbidden]
        public static readonly string ParamArray = "ParamArray";
        [Forbidden]
        public static readonly string Preserve = "Preserve";
        [Forbidden]
        public static readonly string Print = "Print";
        [Forbidden]
        public static readonly string Private = "Private";
        public static readonly string Property = "Property";
        [Forbidden]
        public static readonly string Public = "Public";
        [Forbidden]
        public static readonly string Put = "Put";
        [Forbidden]
        public static readonly string RaiseEvent = "RaiseEvent";
        public static readonly string Random = "Random";
        public static readonly string Randomize = "Randomize";
        public static readonly string Read = "Read";
        [Forbidden]
        public static readonly string ReDim = "ReDim";
        [Forbidden]
        public static readonly string Rem = "Rem";
        [Forbidden]
        public static readonly string Resume = "Resume";
        [Forbidden]
        public static readonly string Return = "Return";
        [Forbidden]
        public static readonly string RSet = "RSet";
        public static readonly string Right = "Right";
        public static readonly string RightB = "RightB";
        public static readonly string RmDir = "RmDir";
        public static readonly string Rnd = "Rnd";
        public static readonly string RTrim = "RTrim";
        public static readonly string Second = "Second";
        [Forbidden]
        public static readonly string Seek = "Seek";
        [Forbidden]
        public static readonly string Select = "Select";
        [Forbidden]
        public static readonly string Set = "Set";
        [Forbidden]
        public static readonly string Shared = "Shared";
        public static readonly string Shell = "Shell";
        public static readonly string Sin = "Sin";
        [Forbidden]
        public static readonly string Single = "Single";
        public static readonly string Sng = "Sng";
        [Forbidden]
        public static readonly string Space = "Space";
        public static readonly string Spc = "Spc";
        public static readonly string Split = "Split";
        public static readonly string Sqr = "Sqr";
        public static readonly string Static = "Static";
        public static readonly string Step = "Step";
        [Forbidden]
        public static readonly string Stop = "Stop";
        public static readonly string Str = "Str";
        public static readonly string StrConv = "StrConv";
        [Forbidden]
        public static readonly string String = "String";
        public static readonly string StrPtr = "StrPtr";
        [Forbidden]
        public static readonly string Sub = "Sub";
        [Forbidden]
        public static readonly string Then = "Then";
        public static readonly string Time = "Time";
        [Forbidden]
        public static readonly string To = "To";
        public static readonly string Trim = "Trim";
        [Forbidden]
        public static readonly string True = "True";
        [Forbidden]
        public static readonly string Type = "Type";
        public static readonly string TypeName = "TypeName";
        [Forbidden]
        public static readonly string TypeOf = "TypeOf";
        [Forbidden]
        public static readonly string UBound = "UBound";
        public static readonly string UCase = "UCase";
        [Forbidden]
        public static readonly string Unlock = "Unlock";
        [Forbidden]
        public static readonly string Until = "Until";
        public static readonly string Val = "Val";
        [Forbidden]
        public static readonly string Variant = "Variant";
        public static readonly string vbBack = "vbBack";
        public static readonly string vbCr = "vbCr";
        public static readonly string vbCrLf = "vbCrLf";
        public static readonly string vbFormFeed = "vbFormFeed";
        public static readonly string vbLf = "vbLf";
        public static readonly string vbNewLine = "vbNewLine";
        public static readonly string vbNullChar = "vbNullChar";
        public static readonly string vbNullString = "vbNullString";
        public static readonly string vbTab = "vbTab";
        public static readonly string vbVerticalTab = "vbVerticalTab";
        public static readonly string WeekDay = "WeekDay";
        [Forbidden]
        public static readonly string Wend = "Wend";
        [Forbidden]
        public static readonly string While = "While";
        public static readonly string Width = "Width";
        [Forbidden]
        public static readonly string With = "With";
        [Forbidden]
        public static readonly string WithEvents = "WithEvents";
        [Forbidden]
        public static readonly string Write = "Write";
        [Forbidden]
        public static readonly string XOr = "Xor";
        public static readonly string Year = "Year";
    }
}
