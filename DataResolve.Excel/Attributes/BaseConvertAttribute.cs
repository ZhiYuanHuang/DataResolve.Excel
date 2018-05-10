using System;
using System.Collections.Generic;
using System.Text;

namespace DataResolve.Excel.Attributes
{
    public abstract class BaseConvertAttribute:Attribute
    {
        public abstract object InvokeConvert(string fieldValue);
    }
}
