namespace Microsoft.Protocols.TestSuites.Common.DataStructures
{
    using System;
    using System.Reflection;

    /// <summary>
    /// This class contains all the properties which are associated with the elements defined in the MS-ASCAL protocol
    /// </summary>
    public class Calendar
    {
        /// <summary>
        /// Gets or sets BusyStatus information of the Calendar
        /// </summary>
        public byte? BusyStatus { get; set; }

        /// <summary>
        /// Gets or sets OrganizerName information of the Calendar
        /// </summary>
        public string OrganizerName { get; set; }

        /// <summary>
        /// Gets or sets OrganizerEmail information of the Calendar
        /// </summary>
        public string OrganizerEmail { get; set; }

        /// <summary>
        /// Gets or sets DtStamp information of the Calendar
        /// </summary>
        public DateTime? DtStamp { get; set; }

        /// <summary>
        /// Gets or sets Sensitivity information of the Calendar
        /// </summary>
        public byte? Sensitivity { get; set; }

        /// <summary>
        /// Gets or sets UID information of the Calendar
        /// </summary>
        public string UID { get; set; }

        /// <summary>
        /// Gets or sets MeetingStatus information of the Calendar
        /// </summary>
        public byte? MeetingStatus { get; set; }

        /// <summary>
        /// Gets or sets Recurrence information of the Calendar
        /// </summary>
        public Response.Recurrence Recurrence { get; set; }

        /// <summary>
        /// Gets or sets Exceptions information of the Calendar
        /// </summary>
        public Response.Exceptions Exceptions { get; set; }

        /// <summary>
        /// Gets or sets ResponseRequested information of the Calendar
        /// </summary>
        public bool? ResponseRequested { get; set; }

        /// <summary>
        /// Gets or sets AppointmentReplyTime information of the Calendar
        /// </summary>
        public DateTime? AppointmentReplyTime { get; set; }

        /// <summary>
        /// Gets or sets ResponseType information of the Calendar
        /// </summary>
        public uint? ResponseType { get; set; }

        /// <summary>
        /// Gets or sets DisallowNewTimeProposal information of the Calendar
        /// </summary>
        public bool? DisallowNewTimeProposal { get; set; }

        /// <summary>
        /// Gets or sets OnlineMeetingConfLink information of the Calendar
        /// </summary>
        public string OnlineMeetingConfLink { get; set; }

        /// <summary>
        /// Gets or sets OnlineMeetingExternalLink information of the Calendar
        /// </summary>
        public string OnlineMeetingExternalLink { get; set; }

        /// <summary>
        /// Gets or sets Reminder information of the Calendar
        /// </summary>
        public uint? Reminder { get; set; }

        /// <summary>
        /// Gets or sets Subject information of the Calendar
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// Gets or sets StartTime information of the Calendar
        /// </summary>
        public DateTime? StartTime { get; set; }

        /// <summary>
        /// Gets or sets EndTime information of the Calendar
        /// </summary>
        public DateTime? EndTime { get; set; }

        /// <summary>
        /// Gets or sets Location information of the Calendar
        /// </summary>
        public string Location { get; set; }

        /// <summary>
        /// Gets or sets Location information of the Calendar
        /// </summary>
        public Response.Location Location1 { get; set; }

        /// <summary>
        /// Gets or sets Attendees information of the Calendar
        /// </summary>
        public Response.Attendees Attendees { get; set; }

        /// <summary>
        /// Gets or sets TimeZone information of the Calendar
        /// </summary>
        public string Timezone { get; set; }

        /// <summary>
        /// Gets or sets AllDayEvent information of the Calendar, the possible value is 0,1
        /// </summary>
        public byte? AllDayEvent { get; set; }

        /// <summary>
        /// Gets or sets Categories information of the Calendar
        /// </summary>
        public Response.Categories Categories { get; set; }

        /// <summary>
        /// Gets or sets NativeBodyType information of the Calendar
        /// </summary>
        public byte? NativeBodyType { get; set; }

        /// <summary>
        /// Gets or sets Calendar body information of the Calendar
        /// </summary>
        public Response.Body Body { get; set; }

        /// <summary>
        /// Deserialize to object instance from SyncCollectionsCollectionCommandsAddApplicationData
        /// </summary>
        /// <typeparam name="T">The generic type parameter</typeparam>
        /// <param name="applicationData">The application data which contains new added information</param>
        /// <returns>The object instance</returns>
        public static T DeserializeFromAddApplicationData<T>(Response.SyncCollectionsCollectionCommandsAddApplicationData applicationData)
        {
            T obj = Activator.CreateInstance<T>();
            if (applicationData.ItemsElementName.Length > 0)
            {
                for (int itemIndex = 0; itemIndex < applicationData.ItemsElementName.Length; itemIndex++)
                {
                    switch (applicationData.ItemsElementName[itemIndex])
                    {
                        case Response.ItemsChoiceType8.Categories1:
                        case Response.ItemsChoiceType8.Categories2:
                        case Response.ItemsChoiceType8.Categories3:
                        case Response.ItemsChoiceType8.Categories4:
                        case Response.ItemsChoiceType8.Recurrence1:
                        case Response.ItemsChoiceType8.Sensitivity1:
                        case Response.ItemsChoiceType8.Subject1:
                        case Response.ItemsChoiceType8.Subject2:
                        case Response.ItemsChoiceType8.Subject3:
                            break;
                        default:
                            SetCalendarPropertyValue(obj, applicationData.ItemsElementName[itemIndex].ToString(), applicationData.Items[itemIndex]);
                            break;
                    }
                }
            }

            return obj;
        }

        /// <summary>
        /// Deserialize to object instance from Properties
        /// </summary>
        /// <typeparam name="T">The generic type parameter</typeparam>
        /// <param name="properties">The Properties data which contains new added information</param>
        /// <returns>The object instance</returns>
        public static T DeserializeFromFetchProperties<T>(Response.Properties properties)
        {
            T obj = Activator.CreateInstance<T>();
            if (properties.ItemsElementName.Length > 0)
            {
                for (int itemIndex = 0; itemIndex < properties.ItemsElementName.Length; itemIndex++)
                {
                    switch (properties.ItemsElementName[itemIndex])
                    {
                        case Response.ItemsChoiceType3.Categories1:
                        case Response.ItemsChoiceType3.Categories2:
                        //case Response.ItemsChoiceType3.Categories3:
                        case Response.ItemsChoiceType3.Categories4:
                        case Response.ItemsChoiceType3.Recurrence1:
                        case Response.ItemsChoiceType3.Sensitivity1:
                        case Response.ItemsChoiceType3.Subject1:
                        //case Response.ItemsChoiceType3.Subject2:
                        case Response.ItemsChoiceType3.Subject3:
                            break;
                        default:
                            SetCalendarPropertyValue(obj, properties.ItemsElementName[itemIndex].ToString(), properties.Items[itemIndex]);
                            break;
                    }
                }
            }

            return obj;
        }

        /// <summary>
        /// Deserialize to object instance from SyncCollectionsCollectionCommandsChangeApplicationData
        /// </summary>
        /// <typeparam name="T">The generic type parameter</typeparam>
        /// <param name="applicationData">The application data which contains new changed information</param>
        /// <returns>The object instance</returns>
        public static T DeserializeFromChangeApplicationData<T>(Response.SyncCollectionsCollectionCommandsChangeApplicationData applicationData)
        {
            T obj = Activator.CreateInstance<T>();
            if (applicationData.ItemsElementName.Length > 0)
            {
                for (int itemIndex = 0; itemIndex < applicationData.ItemsElementName.Length; itemIndex++)
                {
                    switch (applicationData.ItemsElementName[itemIndex])
                    {
                        case Response.ItemsChoiceType7.Categories1:
                        case Response.ItemsChoiceType7.Categories2:
                        case Response.ItemsChoiceType7.Categories3:
                        case Response.ItemsChoiceType7.Categories4:
                        case Response.ItemsChoiceType7.Recurrence1:
                        case Response.ItemsChoiceType7.Sensitivity1:
                        case Response.ItemsChoiceType7.Subject1:
                        case Response.ItemsChoiceType7.Subject2:
                        case Response.ItemsChoiceType7.Subject3:
                            break;
                        default:
                            SetCalendarPropertyValue(obj, applicationData.ItemsElementName[itemIndex].ToString(), applicationData.Items[itemIndex]);
                            break;
                    }
                }
            }

            return obj;
        }

        /// <summary>
        /// Get note instance from the Properties element of Search response
        /// </summary>
        /// <typeparam name="T">The generic type parameter</typeparam>
        /// <param name="properties">The data which contains information for note</param>
        /// <param name="protocolVer">The protocol version specifies the version of ActiveSync protocol used to communicate with the server.</param>
        /// <returns>The returned note instance</returns>
        public static T DeserializeFromSearchProperties<T>(Response.SearchResponseStoreResultProperties properties,string protocolVer)
        {
            T obj = Activator.CreateInstance<T>();

            if (properties.ItemsElementName.Length > 0)
            {
                for (int itemIndex = 0; itemIndex < properties.ItemsElementName.Length; itemIndex++)
                {
                    switch (properties.ItemsElementName[itemIndex])
                    {
                        case Response.ItemsChoiceType6.Categories:
                        case Response.ItemsChoiceType6.Categories1:
                        case Response.ItemsChoiceType6.Categories2:
                        case Response.ItemsChoiceType6.Categories3:
                        case Response.ItemsChoiceType6.Recurrence1:
                        case Response.ItemsChoiceType6.Subject1:
                        case Response.ItemsChoiceType6.Subject2:
                        case Response.ItemsChoiceType6.Subject3:
                        case Response.ItemsChoiceType6.Sensitivity1:
                            break;
                        case Response.ItemsChoiceType6.Location:                         
                            if (protocolVer == "14.0"|| protocolVer == "14.1")
                            {
                                SetCalendarPropertyValue(obj, properties.ItemsElementName[itemIndex].ToString(), properties.Items[itemIndex]);
                            }
                            break;
                        case Response.ItemsChoiceType6.Location1:
                            if (protocolVer == "14.0" || protocolVer == "14.1")
                            {
                                break;
                            }
                            SetCalendarPropertyValue(obj, properties.ItemsElementName[itemIndex].ToString(), properties.Items[itemIndex]);
                            break;
                        default:
                            SetCalendarPropertyValue(obj, properties.ItemsElementName[itemIndex].ToString(), properties.Items[itemIndex]);
                            break;
                    }
                }
            }

            return obj;
        }

        /// <summary>
        /// Set specified property value by property name
        /// </summary>
        /// <param name="targetObject">The object which should be set value</param>
        /// <param name="propertyName">The property name</param>
        /// <param name="value">Property value</param>
        public static void SetCalendarPropertyValue(object targetObject, string propertyName, object value)
        {
            if (string.IsNullOrEmpty(propertyName) || null == value || null == targetObject)
            {
                return;
            }

            Type currentType = targetObject.GetType();
            PropertyInfo property = currentType.GetProperty(propertyName);

            if (property != null)
            {
                if (property.PropertyType == typeof(DateTime?))
                {
                    value = Common.GetNoSeparatorDateTime(value.ToString());
                }
                else if (property.PropertyType == typeof(byte) || property.PropertyType == typeof(byte?))
                {
                    value = byte.Parse(value.ToString());
                }
                else if (property.PropertyType == typeof(bool) || property.PropertyType == typeof(bool?))
                {
                    if (value.ToString() == "0")
                    {
                        value = false;
                    }
                    else if (value.ToString() == "1")
                    {
                        value = true;
                    }
                    else
                    {
                        value = bool.Parse(value.ToString());
                    }
                }
                else if (property.PropertyType == typeof(uint) || property.PropertyType == typeof(uint?))
                {
                    if (!string.IsNullOrEmpty(value.ToString()))
                    {
                        value = uint.Parse(value.ToString());
                    }
                    else
                    {
                        value = null;
                    }
                }
                else if (property.PropertyType == typeof(ushort) || property.PropertyType == typeof(ushort?))
                {
                    value = ushort.Parse(value.ToString());
                }

                property.SetValue(targetObject, value, null);
            }
        }
    }
}