namespace Microsoft.Protocols.TestSuites.MS_OXWSCORE
{
    using System;
    using System.Globalization;
    using System.Reflection;
    using System.Text.RegularExpressions;
    using Common;

    /// <summary>
    /// A class contains all helper methods used in test cases.
    /// </summary>
    public static class TestSuiteHelper
    {
        #region Const strings
        /// <summary>
        /// The subject for creating an item.
        /// </summary>
        public const string SubjectForCreateItem = "Created Item Subject";

        /// <summary>
        /// The body content for an item.
        /// </summary>
        public const string BodyForBaseItem = "This Is The Base Object.";

        /// <summary>
        /// The subject for updating an item.
        /// </summary>
        public const string SubjectForUpdateItem = "Updated Item Subject";

        /// <summary>
        /// An integer that represents a property dispatch identifier which is used as the PropertyId value for ExtendedFieldURI.
        /// </summary>
        public const string PropertyId = "123";

        /// <summary>
        /// A string that represents an extended property by name which is used as the PropertyName value for ExtendedFieldURI.
        /// </summary>
        public const string PropertyName = "Classification";

        /// <summary>
        /// A string for a single MAPI property which is used as the item value for ExtendedProperty.
        /// </summary>
        public const string ElementValue = "12/25/2009 3:25:15 PM";

        /// <summary>
        /// A string value that indicates the message class of the item except IPM.Note, such as IPM.Contact, IPM.Post or IPM.Appointment, which can be referred to MS-OXCFOLD.
        /// </summary>
        public const string ItemClassNotNote = "IPM.Appointment";

        /// <summary>
        /// A string value that indicates an invalid message class.
        /// </summary>
        public const string InvalidItemClass = "IPM.WrongMessage";

        /// <summary>
        /// The number of minutes before an event occurs when a reminder is displayed.
        /// </summary>
        public const string ReminderMinutesBeforeStart = "10";

        /// <summary>
        /// The name of the category for an item.
        /// </summary>
        public const string CategoryName = "The Category";

        /// <summary>
        /// A string value that contains the identifier of the item to which this item is a reply.
        /// </summary>
        public const string InReplyTo = "Someone";

        /// <summary>
        /// The culture for an item in a mailbox.
        /// </summary>
        public const string Culture = "en-US";

        /// <summary>
        /// The ID of a time zone.
        /// </summary>
        public const string TimeZoneID = "Pacific Standard Time";

        /// <summary>
        /// A string element that represents the display name of a contact.
        /// </summary>
        public const string ContactString = "James";
        #endregion

        /// <summary>
        /// Copies an object.
        /// </summary>
        /// <typeparam name="T">The object's type.</typeparam>
        /// <param name="source">Source object.</param>
        /// <returns>The copied object.</returns>
        public static T Copy<T>(object source)
        {
            Type type = typeof(T);
            Assembly assembly = Assembly.GetAssembly(type);
            object destination = assembly.CreateInstance(type.ToString());

            foreach (FieldInfo mi in type.GetFields())
            {
                FieldInfo des = type.GetField(mi.Name);
                if (des != null && des.FieldType == mi.FieldType)
                {
                    des.SetValue(destination, mi.GetValue(source));
                }
            }

            foreach (PropertyInfo pi in type.GetProperties())
            {
                PropertyInfo des = type.GetProperty(pi.Name);
                if (des != null && des.PropertyType == pi.PropertyType && des.CanWrite && pi.CanRead)
                {
                    des.SetValue(destination, pi.GetValue(source, null), null);
                }
            }

            return (T)destination;
        }

        /// <summary>
        /// Get the charset from an HTML string.
        /// </summary>
        /// <param name="html">the html format string.</param>
        /// <returns>The charset string.</returns>
        public static string GetCharsetOfHTML(string html)
        {
            Match charSetMatch = Regex.Match(html, "<meta([^<]*)charset=([^<]*)\"", RegexOptions.IgnoreCase | RegexOptions.Multiline);
            string charSet = charSetMatch.Groups[2].Value;
            return charSet.ToLower(new CultureInfo(TestSuiteHelper.Culture, false));
        }

        /// <summary>
        /// Get the target attribute from an HTML link string.
        /// </summary>
        /// <param name="html">the html format link string.</param>
        /// <returns>The target attribute string.</returns>
        public static string GetTargetAttribute(string html)
        {
            Match targetMatch = Regex.Match(html, "<a([^<]*)target=([^<]*)\"", RegexOptions.IgnoreCase | RegexOptions.Multiline);
            string target = targetMatch.Groups[2].Value;
            return target.ToLower(new CultureInfo(TestSuiteHelper.Culture, false));
        }

        /// <summary>
        /// Whether the body string is HTML format.
        /// </summary>
        /// <param name="body">the body string</param>
        /// <returns>The boolean value indicates whether the body string is HTML format or not.</returns>
        public static bool IsHTML(string body)
        {
            bool htmlMatch = Regex.IsMatch(body, @"<(\S*?)[^>]*>.*?</\1>|<.*? />", RegexOptions.IgnoreCase | RegexOptions.Multiline);
            return htmlMatch;
        }

        /// <summary>
        /// Whether the HTML string contains an image source
        /// </summary>
        /// <param name="html">the html format string</param>
        /// <returns>The boolean value indicates whether the HTML string contains an image source or not.</returns>
        public static bool ContainImageSrcOfHTML(string html)
        {
            bool imageSrcMatch = Regex.IsMatch(html, @"<img\b[^<>]*?\bsrc[\s\t\r\n]*=[\s\t\r\n]*[""']?[\s\t\r\n]*(?<imgUrl>[^\s\t\r\n""'<>]*)[^<>]*?/?[\s\t\r\n]*>", RegexOptions.IgnoreCase | RegexOptions.Multiline);
            return imageSrcMatch;
        }

        /// <summary>
        /// Whether the string is base 64 encoding.
        /// </summary>
        /// <param name="base64String">the string to be checked</param>
        /// <returns>The boolean value indicates whether the string is base 64 encoding or not.</returns>
        public static bool IsBase64String(string base64String)
        {
            bool isBase64 = Regex.IsMatch(base64String, @"^([A-Za-z0-9+/]{4})*([A-Za-z0-9+/]{4}|[A-Za-z0-9+/]{3}=|[A-Za-z0-9+/]{2}==)$", RegexOptions.IgnoreCase | RegexOptions.Multiline);
            return isBase64;
        }

        /// <summary>
        /// Whether the byte array is base 64 binary.
        /// </summary>
        /// <param name="base64Binary">the array of bytes to be checked</param>
        /// <returns>The boolean value indicates whether the byte array is base 64 binary or not.</returns>
        public static bool IsBase64Binary(byte[] base64Binary)
        {
            bool isBase64 = false;
            string base64String = Convert.ToBase64String(base64Binary);
            if (!string.IsNullOrEmpty(base64String))
            {
                isBase64 = true;
            }

            return isBase64;
        }

        /// <summary>
        /// Set the ItemChange element for UpdateItem operation.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="item">The item to be updated.</param>
        /// <param name="index">A parameter that represents the index of the resources of the same type, which is used to combine the unique body.</param>
        /// <returns>Return an ItemChangeType object.</returns>
        public static ItemChangeType CreateItemChangeItem<T>(T item, uint index)
        where T : ItemType, new()
        {
            return new ItemChangeType()
            {
                Item = item.ItemId,

                Updates = new ItemChangeDescriptionType[]
                    {
                        new SetItemFieldType()
                        {
                            Item = new PathToUnindexedFieldType()
                            {
                                FieldURI = UnindexedFieldURIType.itemBody
                            },

                            Item1 = new T()
                            {
                                Body = new BodyType()
                                {
                                    Value = TestSuiteHelper.BodyForBaseItem + index
                                }
                            }
                        }
                    }
            };
        }
    }
}