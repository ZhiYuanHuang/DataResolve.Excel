using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections;
using System.Reflection;
using System.Linq;
using DataResolve.Excel.Attributes;

namespace DataResolve.Excel.Core
{
    public class ExcelConverter : IExcelConvert
    {
        private DataConfig _config;
        public ExcelConverter(DataConfig config)
        {
            _config = config;
        }

        public T DeserializeObject<T>(string filePath)
        {
            ValidInput(filePath);

            FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            IWorkbook workBook = null;
            if (filePath.IndexOf(".xlsx") > 0)
            {
                workBook = new XSSFWorkbook(fileStream);
            }
            else
            {
                workBook = new HSSFWorkbook(fileStream);
            }

            List<string> fieldList = GetFieldListByResolve(workBook);
            List<List<string>> dataList = GetDataListByResolve(workBook);

            Type type = typeof(T);
            object result = null;
            if (type.GetInterface("IList") != null)
            {
                var temp = (IList)Activator.CreateInstance(type);
                Type objType = type.GetGenericArguments()[0];
                var properties = objType.GetProperties();
                foreach (List<string> data in dataList)
                {
                    var obj = Activator.CreateInstance(objType);

                    for (int i = 0; i < data.Count; i++)
                    {
                        string fieldName = fieldList[i];
                        string fieldValue = data[i];
                        SetValue(obj, objType, properties, fieldName, fieldValue);
                    }
                    temp.Add(obj);
                }
                result = temp;
            }
            else
            {
                var temp = Activator.CreateInstance(type);
                var properties = type.GetProperties();
                List<string> data = dataList[0];

                for (int i = 0; i < data.Count; i++)
                {
                    string fieldName = fieldList[i];
                    string fieldValue = data[i];
                    SetValue(temp, type, properties, fieldName, fieldValue);
                }
                result = temp;
            }

            return (T)result;
        }

        private void SetValue(object entity, Type type, PropertyInfo[] properties, string fieldName, string fieldValue)
        {
            foreach (var property in properties)
            {
                if (!property.IsDefined(typeof(DescribeAttribute), false))
                {
                    continue;
                }
                var attribute = property.GetCustomAttribute<DescribeAttribute>();
                if (attribute.FieldName != fieldName)
                {
                    continue;
                }
                object propertyValue = null;
                if (string.IsNullOrEmpty(fieldValue))
                {
                    if (attribute.DefaultValue != null)
                    {
                        propertyValue = attribute.DefaultValue;
                        if (property.IsDefined(typeof(BaseConvertAttribute), false))
                        {
                            var convertAttribute = property.GetCustomAttribute<BaseConvertAttribute>();
                            propertyValue = convertAttribute.InvokeConvert(propertyValue.ToString());
                        }
                        propertyValue = Convert.ChangeType(propertyValue, property.PropertyType);
                    }
                }
                else
                {
                    propertyValue = fieldValue;
                    if (property.IsDefined(typeof(BaseConvertAttribute), false))
                    {
                        var convertAttribute = property.GetCustomAttribute<BaseConvertAttribute>();
                        propertyValue = convertAttribute.InvokeConvert(propertyValue.ToString());
                    }
                    propertyValue = Convert.ChangeType(propertyValue, property.PropertyType);
                }
               
                property.SetValue(entity, propertyValue);
                break;
            }
        }

        public List<string> GetFieldListByResolve(IWorkbook workBook)
        {
            List<string> result = new List<string>();

            ISheet sheet = workBook.GetSheetAt(_config.SheetIndex);
            IRow row = sheet.GetRow(_config.FieldRow);
            for (int i = 0; i < row.LastCellNum; i++)
            {
                ICell cell = row.GetCell(i, MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell == null)
                {
                    continue;
                }
                string cellValue = cell.ToString();
                result.Add(cellValue);
            }

            return result;
        }

        public List<List<string>> GetDataListByResolve(IWorkbook workBook)
        {
            List<List<string>> result = new List<List<string>>();

            ISheet sheet = workBook.GetSheetAt(_config.SheetIndex);
            IRow row;
            IRow fieldRow = sheet.GetRow(_config.FieldRow);
            for (int i = _config.DataStartRow; i <= sheet.LastRowNum; i++)
            {
                row = sheet.GetRow(i);
                if (row != null)
                {
                    List<string> rowValueList = new List<string>();
                    for (int j = 0; j < row.LastCellNum; j++)
                    {
                        ICell fieldCell = fieldRow.GetCell(j, MissingCellPolicy.RETURN_BLANK_AS_NULL);
                        if (fieldCell == null)
                        {
                            continue;
                        }

                        string cellValue = string.Empty;
                        ICell dataCell = row.GetCell(j);
                        if (dataCell != null)
                        {
                            cellValue = dataCell.ToString();
                        }
                        rowValueList.Add(cellValue);
                    }
                    result.Add(rowValueList);
                }
            }

            return result;
        }

        private void ValidInput(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException();
            }
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException();
            }

            if (filePath.IndexOf(".xlsx") < 0 && filePath.IndexOf(".xls") < 0)
            {
                throw new ArgumentException("File format error!");
            }
        }
    }
}
