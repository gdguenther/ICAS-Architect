using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Diagnostics;

namespace ICAS_Architect
{


    internal static class DataTableMethods
    {
        public static DataTable ToDataTable<T>(this System.Collections.Generic.List<T> collection, string _sTableName)
        {
            var table = new DataTable(_sTableName);
            var type = typeof(T);
            var properties = type.GetProperties();

            //Create the columns in the DataTable
            foreach (var property in properties)
            {
                if (property.PropertyType == typeof(int?))
                    table.Columns.Add(property.Name, typeof(int));
                else if (property.PropertyType == typeof(decimal?))
                    table.Columns.Add(property.Name, typeof(decimal));
                else if (property.PropertyType == typeof(double?))
                    table.Columns.Add(property.Name, typeof(double));
                else if (property.PropertyType == typeof(DateTime?))
                    table.Columns.Add(property.Name, typeof(DateTime));
                else if (property.PropertyType == typeof(Guid?))
                    table.Columns.Add(property.Name, typeof(Guid));
                else if (property.PropertyType == typeof(bool?))
                    table.Columns.Add(property.Name, typeof(bool));
                else
                    table.Columns.Add(property.Name, property.PropertyType);
            }

            //Populate the table
            foreach (var item in collection)
            {
                var row = table.NewRow();
                row.BeginEdit();
                foreach (var property in properties)
                    row[property.Name] = property.GetValue(item, null) ?? DBNull.Value;

                row.EndEdit();
                table.Rows.Add(row);
            }

            return table;
        }
//        http://martinwilley.com/net/code/convertdatatable.html
        private static bool CanUseType(Type propertyType)
        {
            //only strings and value types
            if (propertyType.IsArray) return false;
            if (!propertyType.IsValueType && propertyType != typeof(string)) return false;
            return true;
        }

        public static DataTable ConvertToDataTable<T>(IList<T> list, string _sTableName) where T : class
        {
            DataTable table = CreateDataTable<T>();
            table.TableName = _sTableName;
            Type objType = typeof(T);
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(objType);
            //Debug.WriteLine("foreach (" + objType.Name + " item in list) {");
            foreach (T item in list)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor property in properties)
                {
                    if (!CanUseType(property.PropertyType)) continue; //shallow only
                    //Debug.WriteLine("row[\"" + property.Name + "\"] = item." + property.Name + "; //.HasValue ? (object)item." + property.Name + ": DBNull.Value;");
                    row[property.Name] = property.GetValue(item) ?? DBNull.Value; //can't use null
                }
                Debug.WriteLine("//===");
                table.Rows.Add(row);
            }
            return table;
        }

        public static DataTable CreateDataTable<T>() where T : class
        {
            Type objType = typeof(T);
            DataTable table = new DataTable(objType.Name);
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(objType);
            foreach (PropertyDescriptor property in properties)
            {
                Type propertyType = property.PropertyType;
                if (!CanUseType(propertyType)) continue; //shallow only

                //nullables must use underlying types
                if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                    propertyType = Nullable.GetUnderlyingType(propertyType);
                //enums also need special treatment
                if (propertyType.IsEnum)
                    propertyType = Enum.GetUnderlyingType(propertyType); //probably Int32
                //if you have nested application classes, they just get added. Check if this is valid?
                Debug.WriteLine("table.Columns.Add(\"" + property.Name + "\", typeof(" + propertyType.Name + "));");
                table.Columns.Add(property.Name, propertyType);
            }
            return table;
        }

        /// <summary>
        /// Convert DataTable to IList&lt;T&gt;. Some column names should match property names- or you'll have a list of empty T entities.
        /// </summary>
        /// <example><code>
        /// IList&lt;Person&gt; people = new List&lt;Person&gt;
        ///    {
        ///        new Person { Id= 1, DoB = DateTime.Now, Name = "Bob", Sex = Person.Sexes.Male }
        ///    };
        /// DataTable dt = DataUtil.ConvertToDataTable(people);
        /// IList&lt;Person&gt; people2 = DataUtil.ConvertToList&lt;Person&gt;(dt); //round trip
        /// //Note that people2 is a list of cloned Person objects
        /// </code></example>
        public static IList<T> ConvertToList<T>(DataTable dt) where T : class, new()
        {
            if (dt == null || dt.Rows.Count == 0) return null;
            IList<T> list = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T obj = ConvertDataRowToEntity<T>(row);
                list.Add(obj);
            }
            return list;
        }

        /// <summary>
        /// Convert a single DataRow into an object of type T.
        /// </summary>
        public static T ConvertDataRowToEntity<T>(DataRow row) where T : class, new()
        {
            Type objType = typeof(T);
            T obj = Activator.CreateInstance<T>(); //hence the new() contsraint
            //Debug.WriteLine(objType.Name + " = new " + objType.Name + "();");
            foreach (DataColumn column in row.Table.Columns)
            {
                //may error if no match
                PropertyInfo property =
                    objType.GetProperty(column.ColumnName,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                if (property == null || !property.CanWrite)
                {
                    //Debug.WriteLine("//Property " + column.ColumnName + " not in object");
                    continue; //or throw
                }
                object value = row[column.ColumnName];
                if (value == DBNull.Value) value = null;
                property.SetValue(obj, value, null);
                Debug.WriteLine("obj." + property.Name + " = row[\"" + column.ColumnName + "\"];");
            }
            return obj;
        }

    }
}
//    public static class CustomLINQtoDataSetMethods
//    {
//        public static DataTable CopyToDataTable<T>(this IEnumerable<T> source)
//        {
//            return new ObjectShredder<T>().Shred(source, null, null);
//        }

//        public static DataTable CopyToDataTable<T>(this IEnumerable<T> source,
//                                                    DataTable table, LoadOption? options)
//        {
//            return new ObjectShredder<T>().Shred(source, table, options);
//        }
//    }

//    public class ObjectShredder<T>
//    {
//        private System.Reflection.FieldInfo[] _fi;
//        private System.Reflection.PropertyInfo[] _pi;
//        private System.Collections.Generic.Dictionary<string, int> _ordinalMap;
//        private System.Type _type;

//        // ObjectShredder constructor.
//        public ObjectShredder()
//        {
//            _type = typeof(T);
//            _fi = _type.GetFields();
//            _pi = _type.GetProperties();
//            _ordinalMap = new Dictionary<string, int>();
//        }

//        /// <summary>
//        /// Loads a DataTable from a sequence of objects.
//        /// </summary>
//        /// <param name="source">The sequence of objects to load into the DataTable.</param>
//        /// <param name="table">The input table. The schema of the table must match that
//        /// the type T.  If the table is null, a new table is created with a schema
//        /// created from the public properties and fields of the type T.</param>
//        /// <param name="options">Specifies how values from the source sequence will be applied to
//        /// existing rows in the table.</param>
//        /// <returns>A DataTable created from the source sequence.</returns>
//        public DataTable Shred(IEnumerable<T> source, DataTable table, LoadOption? options)
//        {
//            // Load the table from the scalar sequence if T is a primitive type.
//            if (typeof(T).IsPrimitive)
//            {
//                return ShredPrimitive(source, table, options);
//            }

//            // Create a new table if the input table is null.
//            table = table ?? new DataTable(typeof(T).Name);

//            // Initialize the ordinal map and extend the table schema based on type T.
//            table = ExtendTable(table, typeof(T));

//            // Enumerate the source sequence and load the object values into rows.
//            table.BeginLoadData();
//            using (IEnumerator<T> e = source.GetEnumerator())
//            {
//                while (e.MoveNext())
//                {
//                    if (options != null)
//                    {
//                        table.LoadDataRow(ShredObject(table, e.Current), (LoadOption)options);
//                    }
//                    else
//                    {
//                        table.LoadDataRow(ShredObject(table, e.Current), true);
//                    }
//                }
//            }
//            table.EndLoadData();

//            // Return the table.
//            return table;
//        }

//        public DataTable ShredPrimitive(IEnumerable<T> source, DataTable table, LoadOption? options)
//        {
//            // Create a new table if the input table is null.
//            table = table ?? new DataTable(typeof(T).Name);

//            if (!table.Columns.Contains("Value"))
//            {
//                table.Columns.Add("Value", typeof(T));
//            }

//            // Enumerate the source sequence and load the scalar values into rows.
//            table.BeginLoadData();
//            using (IEnumerator<T> e = source.GetEnumerator())
//            {
//                Object[] values = new object[table.Columns.Count];
//                while (e.MoveNext())
//                {
//                    values[table.Columns["Value"].Ordinal] = e.Current;

//                    if (options != null)
//                    {
//                        table.LoadDataRow(values, (LoadOption)options);
//                    }
//                    else
//                    {
//                        table.LoadDataRow(values, true);
//                    }
//                }
//            }
//            table.EndLoadData();

//            // Return the table.
//            return table;
//        }

//        public object[] ShredObject(DataTable table, T instance)
//        {

//            FieldInfo[] fi = _fi;
//            PropertyInfo[] pi = _pi;

//            if (instance.GetType() != typeof(T))
//            {
//                // If the instance is derived from T, extend the table schema
//                // and get the properties and fields.
//                ExtendTable(table, instance.GetType());
//                fi = instance.GetType().GetFields();
//                pi = instance.GetType().GetProperties();
//            }

//            // Add the property and field values of the instance to an array.
//            Object[] values = new object[table.Columns.Count];
//            foreach (FieldInfo f in fi)
//            {
//                values[_ordinalMap[f.Name]] = f.GetValue(instance);
//            }

//            foreach (PropertyInfo p in pi)
//            {
//                values[_ordinalMap[p.Name]] = p.GetValue(instance, null);
//            }

//            // Return the property and field values of the instance.
//            return values;
//        }

//        public DataTable ExtendTable(DataTable table, Type type)
//        {
//            // Extend the table schema if the input table was null or if the value
//            // in the sequence is derived from type T.
//            foreach (FieldInfo f in type.GetFields())
//            {
//                if (!_ordinalMap.ContainsKey(f.Name))
//                {
//                    // Add the field as a column in the table if it doesn't exist
//                    // already.
//                    DataColumn dc = table.Columns.Contains(f.Name) ? table.Columns[f.Name]
//                        : table.Columns.Add(f.Name, f.FieldType);

//                    // Add the field to the ordinal map.
//                    _ordinalMap.Add(f.Name, dc.Ordinal);
//                }
//            }
//            foreach (PropertyInfo p in type.GetProperties())
//            {
//                if (!_ordinalMap.ContainsKey(p.Name))
//                {
//                    // Add the property as a column in the table if it doesn't exist
//                    // already.
//                    DataColumn dc = table.Columns.Contains(p.Name) ? table.Columns[p.Name]
//                        : table.Columns.Add(p.Name, p.PropertyType);

//                    // Add the property to the ordinal map.
//                    _ordinalMap.Add(p.Name, dc.Ordinal);
//                }
//            }

//            // Return the table.
//            return table;
//        }
//    }

//}
