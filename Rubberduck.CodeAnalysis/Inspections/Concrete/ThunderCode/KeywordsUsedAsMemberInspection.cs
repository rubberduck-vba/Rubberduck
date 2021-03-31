using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete.ThunderCode
{
    /// <summary hidden="true">
    /// A ThunderCode inspection that locates instances of various keywords and reserved identifiers used as Type or Enum member names.
    /// </summary>
    /// <why>
    /// This inpection is flagging code we dubbed "ThunderCode", 
    /// code our friend Andrew Jackson would have written to confuse Rubberduck's parser and/or resolver. 
    /// While perfectly legal as Type or Enum member names, these identifiers should be avoided: 
    /// they need to be square-bracketed everywhere they are used.
    /// </why>
    internal sealed class KeywordsUsedAsMemberInspection : DeclarationInspectionBase
    {
        public KeywordsUsedAsMemberInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.EnumerationMember, DeclarationType.UserDefinedTypeMember)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            var normalizedMemberName = declaration.IdentifierName.ToLowerInvariant();
            return ReservedKeywordsInLowerCase.Contains(normalizedMemberName);
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return InspectionResults.KeywordsUsedAsMemberInspection.ThunderCodeFormat(declaration.IdentifierName);
        }

        // MS-VBAL 3.3.5.2 Reserved Identifiers and IDENTIFIER
        private static readonly IEnumerable<string> ReservedKeywords = new []
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
            Tokens.DefBool,
            Tokens.DefByte,
            Tokens.DefCur,
            Tokens.DefDate,
            Tokens.DefDbl,
            Tokens.DefInt,
            Tokens.DefLng,
            Tokens.DefLngLng,
            Tokens.DefLngPtr,
            Tokens.DefObj,
            Tokens.DefSng,
            Tokens.DefStr,
            Tokens.DefVar,
            Tokens.Dim,
            Tokens.Do,
            Tokens.Else,
            Tokens.ElseIf,
            Tokens.End,
            "EndIf",
            Tokens.Enum,
            Tokens.Erase,
            Tokens.Event,
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
            Tokens.Lock,
            Tokens.Loop,
            Tokens.LSet,
            Tokens.Next,
            Tokens.On,
            Tokens.Open,
            Tokens.Option,
            Tokens.Print,
            Tokens.Private,
            Tokens.Public,
            Tokens.Put,
            Tokens.RaiseEvent,
            Tokens.ReDim,
            Tokens.Resume,
            Tokens.Return,
            Tokens.RSet,
            Tokens.Seek,
            Tokens.Select,
            Tokens.Set,
            Tokens.Static,
            Tokens.Stop,
            Tokens.Sub,
            Tokens.Type,
            Tokens.Unlock,
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
            Tokens.Shared,
            Tokens.Until,
            Tokens.WithEvents,
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

        private static readonly HashSet<string> ReservedKeywordsInLowerCase =
            ReservedKeywords.Select(keyword => keyword.ToLowerInvariant())
                .ToHashSet();
    }
}
