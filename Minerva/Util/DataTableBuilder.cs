using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;


namespace Minerva.Util
{
    /// <summary>
    /// DataTable的操作类
    /// </summary>
    class DataTableBuilder
    {
        //将一个List转为一个DataTable，为Word的merge，Cell的生成提供数据源(Aspose不支持List<T>，仅支持DataTable等)
        public static DataTable ToDataTableOfList<T>(List<T> entities,string name)
        {
            if (entities.Count == 0)
            {
                return new DataTable();
            }

            Type type = entities[0].GetType();
            PropertyInfo[] properties = type.GetProperties();
            DataTable dt = new DataTable(name);

            properties.ToList().ForEach(prop =>
            {
                dt.Columns.Add(prop.Name);
            });

            entities.ForEach(entity =>
            {
                object[] values = new object[properties.Length];
                for (int i = 0; i < properties.Length; i++)
                {
                    values[i] = properties[i].GetValue(entity, null);
                }
                dt.Rows.Add(values);
            });

            return dt;
        }
    }
}
