using System;
using System.Collections.Generic;
using System.Text;

namespace DataResolve.Excel.Core
{
    public interface IExcelConvert
    {
        T DeserializeObject<T>(string filePath);
    }
}
