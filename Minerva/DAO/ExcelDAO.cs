using System;
using System.Collections.Generic;
using Aspose.Cells;
using System.Reflection;
using System.Linq;
using System.Text.RegularExpressions;

namespace Minerva.DAO
{
    using Minerva.Util;
    /// <summary>
    /// Excel操作类，用于处理Excel的读写,是一个被封装的Excel处理对象
    /// 每一个类的实例对应一个具体的Excel文件
    /// </summary>
    /// <typeparam name="T"></typeparam>
    class ExcelDAO<T>
    {
        private readonly string excelPath;
        private readonly Workbook workbook;
        private int sheetIndex = 0;
        private int skipRows = 0;


        public ExcelDAO(string excelPath)
        {
            this.excelPath = excelPath;
            workbook = new Workbook(this.excelPath);
        }

        //设置选中的sheet
        public ExcelDAO<T> SelectSheetAt(int sheetIndex)
        {
            this.sheetIndex = sheetIndex;
            return this;
        }

        //设置需要跳过的行数
        public ExcelDAO<T> Skip(int skipRows)
        {
            this.skipRows = skipRows;
            return this;
        }

        //判断是否为数字
        private bool IsDigit(string s)
        {
            return Regex.IsMatch(s, @"^[+-]?\d*$");
        }

        //读取一个sheet，返回对应Type的List
        //是ToEntityList(int index, int skip, Type type)方法的一个重载特例
        public List<T> ToEntityList(Type type)
        {
            return ToEntityList(sheetIndex, skipRows, type);
        }

        //读取一个sheet，返回对应Type的List
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

        //将sheet中的一行转为一个给定Type的对象
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

        //设置一个object对象的指定field的值
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

        //设置一个单元格的值，位于指定sheet的指定坐标[行坐标，列坐标]
        private void SetCellValue(int sheetIndex, int rowIndex, int columnIndex, object value)
        {
            workbook.Worksheets[sheetIndex].Cells[rowIndex, columnIndex].PutValue(value);
        }

        //将一个List写入一个新的sheet
        //是SetCellValues(int index, List<T> entityList)一个重载特例
        public void SetCellValues(List<T> entityList)
        {
            int index = workbook.Worksheets.Add(SheetType.Worksheet);
            this.sheetIndex = index;
            SetCellValues(index, entityList);
        }

        //将一个List写入一个新的sheet
        //是SetCellValues(int sheetIndex, int rowIndex, int columnIndex, List<T> entityList)一个重载特例
        public void SetCellValues(int index, List<T> entityList)
        {
            Worksheet sheet = workbook.Worksheets[index];
            this.sheetIndex = index;
            int rowIndex = sheet.Cells.MaxRow + 1;
            int columnIndex = 0;
            SetCellValues(index, rowIndex, columnIndex, entityList);
        }

        //将一个List写入指定sheet的指定坐标[行坐标，列坐标]
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

        //获取一个对象指定properties的值
        //例如
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

        //自适应Columns的宽度
        public ExcelDAO<T> Fit()
        {
            workbook.Worksheets[this.sheetIndex].AutoFitColumns();
            return this;
        }

        //保存文件，将修改保存到原文件
        public ExcelDAO<T> Save()
        {
            workbook.Save(excelPath);
            return this;
        }

        //另存为，保存至指定新文件路径
        public ExcelDAO<T> SaveAs(string newFilePath)
        {
            workbook.Save(newFilePath);
            return this;
        }
    }
}
