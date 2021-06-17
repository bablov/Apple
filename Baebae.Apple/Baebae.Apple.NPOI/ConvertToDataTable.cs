using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Reflection;

namespace Baebae.Apple.NPOI
{
    public static class ConvertToDataTable
    {
        /// <summary>
        /// List转DataTable
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="list">列表</param>
        /// <param name="header">列头</param>
        /// <returns></returns>
        public static DataTable List2DataTable<T>(this List<T> list, IDictionary<string, string> header = null) where T : class
        {
            //如果header无效
            if (header == null || header.Count == 0)
                return list.GetDataTable(typeof(T));

            DataTable dt = new DataTable();

            PropertyInfo[] p = typeof(T).GetProperties();
            foreach (PropertyInfo pi in p)
            {
                //源数据实体是否包含header列
                if (header.ContainsKey(pi.Name))
                {
                    // The the type of the property
                    Type columnType = pi.PropertyType;

                    // We need to check whether the property is NULLABLE
                    if (pi.PropertyType.IsGenericType && pi.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                    {
                        // If it is NULLABLE, then get the underlying type. eg if "Nullable<int>" then this will return just "int"
                        columnType = pi.PropertyType.GetGenericArguments()[0];
                    }

                    dt.Columns.Add(header[pi.Name], columnType);
                }
            }

            if (list != null)
            {
                for (int i = 0; i < list.Count; i++)
                {
                    IList tempList = new ArrayList();
                    foreach (PropertyInfo pi in p)
                    {
                        object o = pi.GetValue(list[i], null);
                        if (header == null || header.Count == 0 ||  //如果header无效
                            header.ContainsKey(pi.Name))            // 或源数据实体包含header列
                        {
                            tempList.Add(o);
                        }
                    }
                    object[] itm = new object[header.Count];
                    for (int j = 0; j < tempList.Count; j++)
                    {
                        itm.SetValue(tempList[j], j);
                    }
                    dt.LoadDataRow(itm, true);
                }
            }

            return dt;
        }
        /// <summary>
        /// Converts a Generic List into a DataTable
        /// </summary>
        /// <param name="list"></param>
        /// <param name="typ"></param>
        /// <returns></returns>
        internal static DataTable GetDataTable(this IList list, Type typ)
        {
            DataTable dt = new DataTable();

            // Get a list of all the properties on the object
            PropertyInfo[] pi = typ.GetProperties();

            // Loop through each property, and add it as a column to the datatable
            foreach (PropertyInfo p in pi)
            {
                // The the type of the property
                Type columnType = p.PropertyType;

                // We need to check whether the property is NULLABLE
                if (p.PropertyType.IsGenericType && p.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                {
                    // If it is NULLABLE, then get the underlying type. eg if "Nullable<int>" then this will return just "int"
                    columnType = p.PropertyType.GetGenericArguments()[0];
                }

                // Add the column definition to the datatable.
                dt.Columns.Add(new DataColumn(p.Name, columnType));
            }

            // For each object in the list, loop through and add the data to the datatable.
            foreach (object obj in list)
            {
                object[] row = new object[pi.Length];
                int i = 0;

                foreach (PropertyInfo p in pi)
                {
                    row[i++] = p.GetValue(obj, null);
                }

                dt.Rows.Add(row);
            }

            return dt;
        }
    }
}
