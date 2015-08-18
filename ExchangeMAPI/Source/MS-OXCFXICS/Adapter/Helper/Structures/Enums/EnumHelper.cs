namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;

    /// <summary>
    /// Supply help functions for manipulate enumerations.
    /// </summary>
    public class EnumHelper
    {
        /// <summary>
        /// Get all values from an Enumeration.
        /// </summary>
        /// <typeparam name="T">A value type.</typeparam>
        /// <returns>All values of an enumeration.</returns>
        public static List<T> GetEnumValues<T>()
        {
            Type t = typeof(T);
            FieldInfo[] fields = t.GetFields();
            int i;
            List<T> values = new List<T>();

            for (i = 0; i < fields.Length; i++)
            {
                if (fields[i].FieldType == t)
                {
                    values.Add((T)fields[i].GetValue(null));
                }
            }

            return values;
        }
    }
}