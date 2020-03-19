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
        public static string RefactorName = "Move Members, Fields, and/or Constants to another Module";
        public static string Instructions = "Select Declarations and specify a new or existing Destination Module";
        public static string ModuleMatchesProjectNameFailMsg = "Module name matches Project Name";
        public static string SourceAndDestinationModuleNameMatcheFailMsg = "Module name matches Source Module Name";
        public static string MoveMember_Destination = "Destination";
        public static string MoveMember_MemberListLabelFormat = "Source ({0}) Declarations";
        public static string MoveMember_DestinationSelectionLabelFormat = "Destination ({0})";
        public static string MoveMember_SourceModuleLabelFormat => "Source ({0})";

    }
}
