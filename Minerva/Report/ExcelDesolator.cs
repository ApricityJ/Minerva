using System;
using System.Collections.Generic;
using Aspose.Cells;
using System.Reflection;
using System.Linq;

namespace Minerva.Report
{
    class ExcelDesolator<T>
    {
        private string reportPath;
        private Workbook workbook;


        public ExcelDesolator(string reportPath)
        {
            this.reportPath = reportPath;
            workbook = new Workbook(this.reportPath);
        }

        public List<T> ToEntityList(int index, Type type)
        {
            return ToEntityList(index, 0, type);
        }

        public List<T> ToEntityList(int index, int skip, Type type)
        {
            List<T> entityList = new List<T>();
            Worksheet sheet = workbook.Worksheets[index];
            for (int i = skip; i <= sheet.Cells.MaxRow; i++)
            {
                Row row = sheet.Cells.GetRow(i);
                entityList.Add(ToEntity(row, type));
            }
            return entityList;
        }

        private T ToEntity(Row row, Type type)
        {
            PropertyInfo[] properties = type.GetProperties();
            object entity = Activator.CreateInstance(type);
            Dictionary<int, string> classMap = ClassMapper.Instance.ToClassMap(type);
            foreach (KeyValuePair<int, string> pair in classMap)
            {
                string value = row[pair.Key].Value.ToString().Trim();
                string property = pair.Value;
                SetValue(entity, property, value);
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

            if (IsType(propertyInfo.PropertyType, "Minerva.Report.Division") ||
                IsType(propertyInfo.PropertyType, "Minerva.Report.TaskType"))
            {
                if (fieldValue != "")
                    propertyInfo.SetValue(entity, Enum.Parse(propertyInfo.PropertyType, fieldValue), null);
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

        public void SetCellValues(int index, List<T> entityList)
        {
            Worksheet sheet = workbook.Worksheets[index];
            int rowIndex = sheet.Cells.MaxRow + 1;
            int columnIndex = 0;
            SetCellValues(index, rowIndex, columnIndex, entityList);
        }

        public void SetCellValues(int sheetIndex, int rowIndex, int columnIndex, List<T> entityList)
        {

            object o, value;
            PropertyInfo[] properties;
            for (int i = 0; i < entityList.Count; i++)
            {
                o = entityList[i];
                properties = o.GetType().GetProperties();
                for (int j = 0; j < properties.Length; j++)
                {
                    value = ToObjectValue(properties[j].Name, o);
                    SetCellValue(sheetIndex, rowIndex + i, columnIndex + j, value);
                }
            }

        }

        private object ToObjectValue(string name, object obj)
        {
            object value = obj.GetType().GetRuntimeProperties()
                .Where(p => p.Name == name)
                .First()
                .GetValue(obj, null);

            if (value.GetType().GetTypeInfo().BaseType == typeof(Enum))
            {
                return Translator.Instance.ToValueForHuman(value);
            }

            if (value.GetType() == typeof(DateTime))
            {
                return ((DateTime)value).ToString("D");
            }
            return value;

        }
    }
}
