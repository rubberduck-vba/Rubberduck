using Rubberduck.Refactorings.EncapsulateField;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    public class MasterDetailSelectionManager
    {
        private const string _neverATargetID = "_Never_a_TargetID_";
        private bool _detailFieldIsFlagged;

        public MasterDetailSelectionManager(IEncapsulateFieldCandidate selected)
            : this(selected?.TargetID)
        {
            if (selected != null)
            {
                DetailField = new ViewableEncapsulatedField(selected);
            }
        }

        private MasterDetailSelectionManager(string targetID)
        {
            SelectionTargetID = targetID;
            DetailField = null;
            _detailFieldIsFlagged = false;
        }


        private IEncapsulatedFieldViewData _detailField;
        public IEncapsulatedFieldViewData DetailField
        {
            set
            {
                _detailField = value;
                _detailFieldIsFlagged = _detailField?.EncapsulateFlag ?? false;
            }
            get => _detailField;
        }

        private string _selectionTargetID;
        public string SelectionTargetID
        {
            set => _selectionTargetID = value;
            get => _selectionTargetID ?? _neverATargetID;
        }

        public bool DetailUpdateRequired
        {
            get
            {
                if (DetailField is null)
                {
                    return true;
                }

                if (_detailFieldIsFlagged != DetailField.EncapsulateFlag)
                {
                    _detailFieldIsFlagged = !_detailFieldIsFlagged;
                    return true;
                }
                return SelectionTargetID != DetailField?.TargetID;
            }
        }
    }

}
