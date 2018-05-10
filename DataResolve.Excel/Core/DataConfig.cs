using System;
using System.Collections.Generic;
using System.Text;

namespace DataResolve.Excel.Core
{
    public class DataConfig
    {
        public DataConfig()
        {
            FieldRow = 0;
            DataStartRow = 1;
            SheetIndex = 0;
        }

        /// <summary>
        /// 字段行
        /// </summary>
        public int FieldRow { get; set; }

        /// <summary>
        /// 数据开始行
        /// </summary>
        public int DataStartRow { get; set; }

        /// <summary>
        /// sheet索引
        /// </summary>
        public int SheetIndex { get; set; }
    }
}
