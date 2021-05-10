using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Minerva.Report
{
    class DataTableUtil
    {
        public static DataTable ToDataTableOfList<T>(List<T> entities,string name)
        {
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
