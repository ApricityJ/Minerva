using System;
using System.Collections.Generic;
using Aspose.Cells;
using System.Reflection;
using System.Linq;
using System.Text.RegularExpressions;
using System.Data;

namespace Minerva.DAO
{
    using Minerva.Util;
    using System.IO;

    /// <summary>
    /// Excel操作类，用于处理Excel的读写,是一个被封装的Excel处理对象
    /// 每一个类的实例对应一个具体的Excel文件
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ExcelDAO<T>
    {
        private readonly string excelPath;
        private readonly Workbook workbook;
        private WorkbookDesigner designer;
        private int sheetIndex = 0;
        private int skipRows = 0;


        public ExcelDAO(string excelPath) 
        {
            this.excelPath = excelPath;

            if (File.Exists(excelPath))
            {
                workbook = new Workbook(this.excelPath);
            }
            else
            {
                workbook = new Workbook();
            }
                
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
            for (int i = skip,length = sheet.Cells.MaxDataRow; i <= length ; i++)
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

        public void SetCellValue(int rowIndex, int columnIndex, object value)
        {
            workbook.Worksheets[sheetIndex].Cells[rowIndex, columnIndex].PutValue(value);
        }

        //设置一个单元格的值，位于指定sheet的指定坐标[行坐标，列坐标]
        public void SetCellValue(int sheetIndex, int rowIndex, int columnIndex, object value)
        {
            workbook.Worksheets[sheetIndex].Cells[rowIndex, columnIndex].PutValue(value);
        }

        public ExcelDAO<T> Name(string name)
        {
            Worksheet sheet = workbook.Worksheets[this.sheetIndex];
            sheet.Name = name;
            return this;
        }

        public ExcelDAO<T> RemoveAt(string name)
        {
            workbook.Worksheets.RemoveAt(name);
            return this;
        }

        //将一个List写入一个新的sheet
        //是SetCellValues(int index, List<T> entityList)一个重载特例
        public ExcelDAO<T> SetCellValues(List<T> entityList)
        {
            int index = workbook.Worksheets.Add(SheetType.Worksheet);
            this.sheetIndex = index;
            SetCellValues(index, entityList);
            return this;
        }

        //将一个List写入一个新的sheet
        //是SetCellValues(int sheetIndex, int rowIndex, int columnIndex, List<T> entityList)一个重载特例
        public ExcelDAO<T> SetCellValues(int index, List<T> entityList)
        {
            Worksheet sheet = workbook.Worksheets[index];
            this.sheetIndex = index;
            int rowIndex = sheet.Cells.MaxDataRow + 1;
            int columnIndex = 0;
            SetCellValues(index, rowIndex, columnIndex, entityList);
            return this;
        }

        //将一个List写入指定sheet的指定坐标[行坐标，列坐标]
        public ExcelDAO<T> SetCellValues(int sheetIndex, int rowIndex, int columnIndex, List<T> entityList)
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
            return this;
        }

        //获取一个对象指定properties的值
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

        public int FindRowIndex(int columnIndex, string s)
        {
            Worksheet sheet = workbook.Worksheets[sheetIndex];
            for (int i = 0; i <= sheet.Cells.MaxDataRow; i++)
            {
                Row row = sheet.Cells.GetRow(i);
                if (row[columnIndex].GetStringValue(CellValueFormatStrategy.CellStyle).Equals(s))
                {
                    return i;
                }
            }

            return -1;
        }

        public int[,] FindCoordinate(string s)
        {
            Worksheet sheet = workbook.Worksheets[sheetIndex];
            for (int i = 0, rows = sheet.Cells.MaxDataRow; i <= rows; i++)
            {
                Row row = sheet.Cells.GetRow(i);
                for (int j = 0, columns = sheet.Cells.MaxDataColumn; j < columns; j++)
                {
                    if (row[j].GetStringValue(CellValueFormatStrategy.CellStyle).Equals(s))
                    {
                        return new int[2, 1] { { i }, { j } };
                    }
                }
            }
            return new int[2, 1] { { -1 }, { -1 } };
        }

        //打开一个设计器
       public ExcelDAO<T> ToDesigner()
        {
            designer = new WorkbookDesigner(workbook);
            return this;
        }

        public ExcelDAO<T> SetDataSource(Dictionary<string,object> dataSet)
        {
            dataSet.ToList().ForEach(p =>
            {
                SetDataSource(p.Key, p.Value);
            });
            return this;
        }

        public ExcelDAO<T> SetDataSource(string k,object v)
        {
            if (v.GetType() == typeof(DataTable))
            {
                designer.SetDataSource((DataTable)v);
            }
            else
            {
                designer.SetDataSource(k, v);
            }
            return this;
        }

        public ExcelDAO<T> Process()
        {
            designer.Process();
            return this;
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

        public void Close()
        {
            this.workbook.Dispose();
        }
    }
}
