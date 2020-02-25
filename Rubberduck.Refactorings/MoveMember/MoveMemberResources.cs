using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    //TODO: These will be assigned keys in RubberduckUI.resx before the WIP PR tag is removed
    public class MoveMemberResources
    {
        public static string Caption => "Move Member";
        public static string ApplicableStrategyNotFound => "Applicable move strategy not found";
        public static string UnsupportedMoveExceptionFormat => "Unable to Move Member: {0}";
        public static string UndefinedDestinationModule => "Undefined Destination Module";
        public static string NoDeclarationsSelectedToMove => "No Declarations Selected to Move";
        public static string MovedContentBelowThisLine => "Moved Content below this line";
        public static string MovedContentAboveThisLine => "Moved Content above this line";
    }
}
