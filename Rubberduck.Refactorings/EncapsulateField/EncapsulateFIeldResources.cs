using Rubberduck.Resources;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    //TODO: All the raw strings here will be assigned keys in RubberduckUI.resx 
    //just before the 'WIP' annotation is removed
    public class EncapsulateFieldResources
    {
        public static string PreviewMarker
            => "'<===== Property and Declaration changes above this line =====>";

        public static string DefaultPropertyParameter => "value";

        public static string DefaultStateUDTFieldName => "this";

        public static string GroupBoxHeaderSuffix = "Property Name:";

        public static string Caption
            => RubberduckUI.EncapsulateField_Caption;

        public static string InstructionText  /* => RubberduckUI.EncapsulateField_InstructionText*/
            => "Select one or more fields to encapsulate.  Accept the default values or edit property names";

        public static string Preview
            => RubberduckUI.EncapsulateField_Preview;

        public static string TitleText
            => RubberduckUI.EncapsulateField_TitleText;

        public static string PrivateUDTPropertyText
            => "Creates a Property for Each UDT Member";

        public static string NameConflictDetected => "Name Conflict Detected";

        public static string ArrayHasExternalRedimFormat => "Storage space for {0} is reallocated external to module '{1}'";
    }
}
