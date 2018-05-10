using System;
using System.Collections.Generic;
using System.Text;

namespace DataResolve.Excel.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class DescribeAttribute:Attribute
    {
        public DescribeAttribute(string fieldName)
        {
            FieldName = fieldName;
        }


        public DescribeAttribute(string fieldName,object defaultValue)
        {
            FieldName = fieldName;
            DefaultValue = defaultValue;
        }

        /// <summary>
        /// 字段名
        /// </summary>
        public string FieldName { get; set; }

        /// <summary>
        /// 缺失或null的默认值
        /// </summary>
        public object DefaultValue { get; set; }
        
    }
}
