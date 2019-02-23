using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Inspections.Inspections.Concrete.ThunderCode
{
    public class KeywordsUsedAsMemberInspection : InspectionBase
    {
        public KeywordsUsedAsMemberInspection(RubberduckParserState state) : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return State.DeclarationFinder.UserDeclarations(DeclarationType.UserDefinedTypeMember)
                .Concat(State.DeclarationFinder.UserDeclarations(DeclarationType.EnumerationMember))
                .Where(m => ReservedKeywords.Any(k => 
                    k.ToLowerInvariant().Equals(
                        m.IdentifierName.Trim().TrimStart('[').TrimEnd(']').ToLowerInvariant())))
                .Select(m => new DeclarationInspectionResult(
                    this,
                    InspectionResults.KeywordsUsedAsMemberInspection.
                        ThunderCodeFormat(m.IdentifierName),
                    m
                ));
        }

        // MS-VBAL 3.3.5.2 Reserved Identifiers and IDENTIFIER
        private static IEnumerable<string> ReservedKeywords = new []
        {
            /*
Statement-keyword = "Call" / "Case" /"Close" / "Const"/ "Declare" / "DefBool" / "DefByte" / 
                    "DefCur" / "DefDate" / "DefDbl" / "DefInt" / "DefLng" / "DefLngLng" / 
                    "DefLngPtr" / "DefObj" / "DefSng" / "DefStr" / "DefVar" / "Dim" / "Do" / 
                    "Else" / "ElseIf" / "End" / "EndIf" /  "Enum" / "Erase" / "Event" / 
                    "Exit" / "For" / "Friend" / "Function" / "Get" / "Global" / "GoSub" / 
                    "GoTo" / "If" / "Implements"/ "Input" / "Let" / "Lock" / "Loop" / 
                    "LSet"/ "Next" / "On" / "Open" / "Option" / "Print" / "Private" / 
                    "Public" / "Put" / "RaiseEvent" / "ReDim" / "Resume" / "Return" / 
                    "RSet" / "Seek" / "Select" / "Set" / "Static" / "Stop" / "Sub" / 
                    "Type" / "Unlock" / "Wend" / "While" / "With" / "Write" 
                    */
                    
            Tokens.Call,
            Tokens.Case,
            Tokens.Close,
            Tokens.Const,
            Tokens.Declare,
            "DefBool",
            "DefByte",
            "DefCur",
            "DefDate",
            "DefDbl",
            "DefInt",
            "DefLng",
            "DefLngLng",
            "DefLngPtr",
            "DefObj",
            "DefSng",
            "DefStr",
            "DefVar",
            Tokens.Dim,
            Tokens.Do,
            Tokens.Else,
            Tokens.ElseIf,
            Tokens.End,
            "EndIf",
            Tokens.Enum,
            "Erase",
            "Event",
            Tokens.Exit,
            Tokens.For,
            Tokens.Friend,
            Tokens.Function,
            Tokens.Get,
            Tokens.Global,
            Tokens.GoSub,
            Tokens.GoTo,
            Tokens.If,
            Tokens.Implements,
            Tokens.Input,
            Tokens.Let,
            "Lock",
            Tokens.Loop,
            "LSet",
            Tokens.Next,
            Tokens.On,
            Tokens.Open,
            Tokens.Option,
            Tokens.Print,
            Tokens.Private,
            Tokens.Public,
            Tokens.Put,
            "RaiseEvent",
            Tokens.ReDim,
            Tokens.Resume,
            Tokens.Return,
            "RSet",
            "Seek",
            Tokens.Select,
            Tokens.Set,
            Tokens.Static,
            Tokens.Stop,
            Tokens.Sub,
            Tokens.Type,
            "Unlock",
            Tokens.Wend,
            Tokens.While,
            Tokens.With,
            Tokens.Write,

            /*
rem-keyword = "Rem" marker-keyword = "Any" / "As"/ "ByRef" / "ByVal "/"Case" / "Each" / 
              "Else" /"In"/ "New" / "Shared" / "Until" / "WithEvents" / "Write" / "Optional" / 
              "ParamArray" / "Preserve" / "Spc" / "Tab" / "Then" / "To" 
            */
            
            Tokens.Any,
            Tokens.As,
            Tokens.ByRef,
            Tokens.ByVal,
            Tokens.Case,
            Tokens.Each,
            Tokens.In,
            Tokens.New,
            "Shared",
            Tokens.Until,
            "WithEvents",
            Tokens.Optional,
            Tokens.ParamArray,
            Tokens.Preserve,
            Tokens.Spc,
            "Tab",
            Tokens.Then,
            Tokens.To,

            /*
operator-identifier = "AddressOf" / "And" / "Eqv" / "Imp" / "Is" / "Like" / "New" / "Mod" / 
                      "Not" / "Or" / "TypeOf" / "Xor"
             */
            
            Tokens.AddressOf,
            Tokens.And,
            Tokens.Eqv,
            Tokens.Imp,
            Tokens.Is,
            Tokens.Like,
            Tokens.Mod,
            Tokens.Not,
            Tokens.Or,
            Tokens.TypeOf,
            Tokens.XOr
        };
    }
}
