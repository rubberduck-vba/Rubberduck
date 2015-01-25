using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.SourceControl
{
    public enum ActionStatus { success, failure }
    public class ActionResult
    {
        public ActionStatus Status { get; private set; }
        public string Message { get; private set; }

        public ActionResult(ActionStatus status, string message)
        {
            this.Status = status;
            this.Message = message;
        }
    }
}
