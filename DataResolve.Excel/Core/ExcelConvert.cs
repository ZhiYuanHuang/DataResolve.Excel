using System;
using System.Collections.Generic;
using System.Text;

namespace DataResolve.Excel.Core
{
    public class ExcelConvert
    {
        public static T DeserializeObject<T>(string filePath, DataConfig config = null)
        {
            if (config == null)
            {
                config = new DataConfig();
            }
            ExcelConverter converter = new ExcelConverter(config);
            return converter.DeserializeObject<T>(filePath);
        }
    }
}
