﻿using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.ComCommands
{
    [ComVisible(false)]
    public class NoIndentAnnotationCommand : ComCommandBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;

        public NoIndentAnnotationCommand(
            IVBE vbe, 
            RubberduckParserState state, 
            IVbeEvents vbeEvents)
            : base(vbeEvents)
        {
            _vbe = vbe;
            _state = state;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            var target = FindTarget(parameter);
            using (var pane = _vbe.ActiveCodePane)
            {
                return pane != null 
                       && !pane.IsWrappingNullReference 
                       && target != null 
                       && !target.Annotations.Any(a => a is NoIndentAnnotation);
            }
        }

        protected override void OnExecute(object parameter)
        {
            using (var activePane = _vbe.ActiveCodePane)
            {
                if (activePane == null || activePane.IsWrappingNullReference)
                {
                    return;
                }

                using (var codeModule = activePane.CodeModule)
                {
                    codeModule.InsertLines(1, "'@NoIndent");
                }
            }
        }

        private Declaration FindTarget(object parameter)
        {
            if (parameter is Declaration declaration)
            {
                return declaration;
            }

            Declaration selectedDeclaration;
            using (var activePane = _vbe.ActiveCodePane)
            {
                selectedDeclaration = _state.FindSelectedDeclaration(activePane);
            }

            while (selectedDeclaration != null && selectedDeclaration.DeclarationType.HasFlag(DeclarationType.Module))
            {
                selectedDeclaration = selectedDeclaration.ParentDeclaration;
            }

            return selectedDeclaration;
        }
    }
}
