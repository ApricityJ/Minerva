using System;
using System.Collections.Generic;
using Aspose.Cells;
using System.Reflection;
using System.Linq;
using System.Text.RegularExpressions;

namespace Minerva.Report
{
    class ExcelDesolator<T>
    {
        private readonly string excelPath;
        private readonly Workbook workbook;
        private int sheetIndex = 0;
        private int skipRows = 0;


        public ExcelDesolator(string excelPath)
        {
            this.excelPath = excelPath;
            workbook = new Workbook(this.excelPath);
        }

        public ExcelDesolator<T> SelectSheetAt(int sheetIndex)
        {
            this.sheetIndex = sheetIndex;
            return this;
        }

        public ExcelDesolator<T> Skip(int skipRows)
        {
            this.skipRows = skipRows;
            return this;
        }

        public List<T> ToEntityList(Type type)
        {
            return ToEntityList(sheetIndex, skipRows, type);
        }

        private bool IsDigit(string s)
        {
            return Regex.IsMatch(s, @"^[+-]?\d*$");
        }

        public List<T> ToEntityList(int index, int skip, Type type)
        {
            List<T> entityList = new List<T>();
            
            Worksheet sheet = workbook.Worksheets[index];
            for (int i = skip; i <= sheet.Cells.MaxRow; i++)
            {
                Row row = sheet.Cells.GetRow(i);
                //exit when first cell value is NOT numeric
                //dangerous but effective,still acceptable
                if (IsDigit(row[0].StringValue))
                {
                    entityList.Add(ToEntity(row, type));
                }
                else
                {
                    break;
                }
            }
            return entityList;
        }

        private T ToEntity(Row row, Type type)
        {
            object entity = Activator.CreateInstance(type);
            Dictionary<int, string> classMap = ClassMapper.Instance.ToClassMap(type);
            foreach (KeyValuePair<int, string> pair in classMap)
            {
                object value = row[pair.Key].Value;
                if (null == value)
                {
                    value = "";
                }
                string property = pair.Value;
                SetValue(entity, property, value.ToString().Trim());
            }
            return (T)entity;
        }

        private void SetValue(object entity, string fieldName, string fieldValue)
        {
            Type type = entity.GetType();
            PropertyInfo propertyInfo = type.GetProperty(fieldName);

            if (IsType(propertyInfo.PropertyType, "System.String"))
            {
                propertyInfo.SetValue(entity, fieldValue, null);
            }

            if (IsType(propertyInfo.PropertyType, "System.Boolean"))
            {
                propertyInfo.SetValue(entity, Boolean.Parse(fieldValue), null);
            }

            if (IsType(propertyInfo.PropertyType, "System.Int"))
            {
                if (fieldValue != "")
                    propertyInfo.SetValue(entity, int.Parse(fieldValue), null);
                else
                    propertyInfo.SetValue(entity, 0, null);
            }
        }

        private bool IsType(Type type, string typeName)
        {
            if (type.ToString() == typeName)
                return true;
            if (type.ToString() == "System.Object")
                return false;
            return IsType(type.BaseType, typeName);
        }

        private void SetCellValue(int sheetIndex, int rowIndex, int columnIndex, object value)
        {
            workbook.Worksheets[sheetIndex].Cells[rowIndex, columnIndex].PutValue(value);
        }

        public void SetCellValues(List<T> entityList)
        {
            int index = workbook.Worksheets.Add(SheetType.Worksheet);
            this.sheetIndex = index;
            SetCellValues(index, entityList);
        }

        public void SetCellValues(int index, List<T> entityList)
        {
            Worksheet sheet = workbook.Worksheets[index];
            this.sheetIndex = index;
            int rowIndex = sheet.Cells.MaxRow + 1;
            int columnIndex = 0;
            SetCellValues(index, rowIndex, columnIndex, entityList);
        }

        public void SetCellValues(int sheetIndex, int rowIndex, int columnIndex, List<T> entityList)
        {
            this.sheetIndex = sheetIndex;
            object entity, value;
            PropertyInfo[] properties;
            for (int i = 0; i < entityList.Count; i++)
            {
                entity = entityList[i];
                properties = entity.GetType().GetProperties();
                for (int j = 0; j < properties.Length; j++)
                {
                    value = ToObjectValue(properties[j].Name, entity);
                    SetCellValue(sheetIndex, rowIndex + i, columnIndex + j, value);
                }
            }
        }

        private object ToObjectValue(string name, object entity)
        {
            object value = entity.GetType().GetRuntimeProperties()
                .Where(p => p.Name == name)
                .First()
                .GetValue(entity, null);

            if (null == value)
            {
                return "";
            }

            if (value.GetType() == typeof(DateTime))
            {
                return ((DateTime)value).ToString("D");
            }
            return value;

        }

        public ExcelDesolator<T> Fit()
        {
            workbook.Worksheets[this.sheetIndex].AutoFitColumns();
            return this;
        }

        public ExcelDesolator<T> Save()
        {
            workbook.Save(excelPath);
            return this;
        }

        public ExcelDesolator<T> SaveAs(string newFilePath)
        {
            workbook.Save(newFilePath);
            return this;
        }
    }
}
