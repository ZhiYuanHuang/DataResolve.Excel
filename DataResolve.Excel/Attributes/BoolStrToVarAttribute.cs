using System;
using System.Collections.Generic;
using System.Text;

namespace DataResolve.Excel.Attributes
{
    public class BoolStrToVarAttribute : BaseConvertAttribute
    {
        public BoolStrToVarAttribute(string trueStr,string falseStr)
        {
            TrueStr = trueStr;
            FalseStr = falseStr;
        }

        public string TrueStr { get; set; }
        public string FalseStr { get; set; }

        public override object InvokeConvert(string fieldValue)
        {
            bool? result = null;

            if (fieldValue == TrueStr)
            {
                result = true;
            }
            else if (fieldValue == FalseStr)
            {
                result = false;
            }

            return result;
        }
    }
}
