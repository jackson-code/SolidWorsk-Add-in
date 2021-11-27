using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWCSharpAddin.ToleranceNetwork.TNData
{
    class TnMateConstraint: TnConstraint
    {
        public TnMateType_e MateType
        { get; private set; }
        public bool UserError = false;

        public TnMateConstraint(TnMateType_e mateType)
        {
            MateType = mateType;
        }

        public override string ToString()
        {
            string result = MateType.ToString();
            if (UserError)
            {
                result += "(UserError)";
            }
            return result;
        }

    }
}
