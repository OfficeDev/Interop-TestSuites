namespace Microsoft.Protocols.TestSuites.Common.DataStructures
{
    using System;

    /// <summary>
    /// This class contains all the properties which are associated with the elements defined in the MS-ASNOTE protocol
    /// </summary>
    public class Note
    {
        /// <summary>
        /// Gets or sets note body information of the note, which contains type and content of the note
        /// </summary>
        public Response.Body Body { get; set; }

        /// <summary>
        /// Gets or sets Categories information of the note
        /// </summary>
        public Response.Categories4 Categories { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether include LastModifiedDate in the response 
        /// </summary>
        public bool IsLastModifiedDateSpecified { get; set; }

        /// <summary>
        /// Gets or sets LastModifiedDate information of the note
        /// </summary>
        public DateTime LastModifiedDate { get; set; }

        /// <summary>
        /// Gets or sets MessageClass information of the note
        /// </summary>
        public string MessageClass { get; set; }

        /// <summary>
        /// Gets or sets subject information of the note
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// Gets or sets the value of LastModifiedDate element from the server response
        /// </summary>
        public string LastModifiedDateString { get; set; }

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
                        case Response.ItemsChoiceType3.Categories:
                        case Response.ItemsChoiceType3.Categories1:
                        case Response.ItemsChoiceType3.Categories2:
                        case Response.ItemsChoiceType3.Categories4:
                        case Response.ItemsChoiceType3.LastModifiedDate:
                        case Response.ItemsChoiceType3.Subject:
                        case Response.ItemsChoiceType3.Subject1:
                        case Response.ItemsChoiceType3.Subject3:
                        case Response.ItemsChoiceType3.MessageClass:
                            break;
                        case Response.ItemsChoiceType3.Categories3:
                            Common.SetSpecifiedPropertyValueByName(obj, "Categories", properties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType3.Body:
                            Common.SetSpecifiedPropertyValueByName(obj, "Body", properties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType3.LastModifiedDate1:
                            string lastModifiedDateString = properties.Items[itemIndex].ToString();
                            if (!string.IsNullOrEmpty(lastModifiedDateString))
                            {
                                Common.SetSpecifiedPropertyValueByName(obj, "IsLastModifiedDateSpecified", true);
                                Common.SetSpecifiedPropertyValueByName(obj, "LastModifiedDate", Common.GetNoSeparatorDateTime(lastModifiedDateString));
                                Common.SetSpecifiedPropertyValueByName(obj, "LastModifiedDateString", lastModifiedDateString);
                            }

                            break;
                        case Response.ItemsChoiceType3.Subject2:
                            Common.SetSpecifiedPropertyValueByName(obj, "Subject", properties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType3.MessageClass1:
                            Common.SetSpecifiedPropertyValueByName(obj, "MessageClass", properties.Items[itemIndex]);
                            break;
                        default:
                            Common.SetSpecifiedPropertyValueByName(obj, properties.ItemsElementName[itemIndex].ToString(), properties.Items[itemIndex]);
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
        /// <returns>The returned note instance</returns>
        public static T DeserializeFromSearchProperties<T>(Response.SearchResponseStoreResultProperties properties)
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
                        case Response.ItemsChoiceType6.Categories3:
                        case Response.ItemsChoiceType6.Subject:
                        case Response.ItemsChoiceType6.Subject1:
                        case Response.ItemsChoiceType6.Subject3:
                        case Response.ItemsChoiceType6.LastModifiedDate:
                        case Response.ItemsChoiceType6.MessageClass:
                            break;
                        case Response.ItemsChoiceType6.Categories2:
                            Common.SetSpecifiedPropertyValueByName(obj, "Categories", (Response.Categories4)properties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType6.Subject2:
                            Common.SetSpecifiedPropertyValueByName(obj, "Subject", properties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType6.Body:
                            Common.SetSpecifiedPropertyValueByName(obj, "Body", (Response.Body)properties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType6.LastModifiedDate1:
                            string lastModifiedDateString = properties.Items[itemIndex].ToString();
                            if (!string.IsNullOrEmpty(lastModifiedDateString))
                            {
                                Common.SetSpecifiedPropertyValueByName(obj, "IsLastModifiedDateSpecified", true);
                                Common.SetSpecifiedPropertyValueByName(obj, "LastModifiedDate", Common.GetNoSeparatorDateTime(lastModifiedDateString));
                                Common.SetSpecifiedPropertyValueByName(obj, "LastModifiedDateString", lastModifiedDateString);
                            }

                            break;
                        case Response.ItemsChoiceType6.MessageClass1:
                            Common.SetSpecifiedPropertyValueByName(obj, "MessageClass", properties.Items[itemIndex]);
                            break;
                        default:
                            Common.SetSpecifiedPropertyValueByName(obj, properties.ItemsElementName[itemIndex].ToString(), properties.Items[itemIndex]);
                            break;
                    }
                }
            }

            return obj;
        }

        /// <summary>
        /// Get note instance from the ApplicationData element of Sync add response
        /// </summary>
        /// <typeparam name="T">The generic type parameter</typeparam>
        /// <param name="applicationData">The data which contains information for note</param>
        /// <returns>The returned instance</returns>
        public static T DeserializeFromAddApplicationData<T>(Response.SyncCollectionsCollectionCommandsAddApplicationData applicationData)
        {
            T obj = Activator.CreateInstance<T>();

            if (applicationData.ItemsElementName.Length > 0)
            {
                for (int itemIndex = 0; itemIndex < applicationData.ItemsElementName.Length; itemIndex++)
                {
                    switch (applicationData.ItemsElementName[itemIndex])
                    {
                        case Response.ItemsChoiceType8.Categories:
                        case Response.ItemsChoiceType8.Categories1:
                        case Response.ItemsChoiceType8.Categories2:
                        case Response.ItemsChoiceType8.Categories4:
                        case Response.ItemsChoiceType8.Subject:
                        case Response.ItemsChoiceType8.Subject1:
                        case Response.ItemsChoiceType8.Subject3:
                        case Response.ItemsChoiceType8.LastModifiedDate:
                        case Response.ItemsChoiceType8.MessageClass:
                            break;
                        case Response.ItemsChoiceType8.Categories3:
                            Common.SetSpecifiedPropertyValueByName(obj, "Categories", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType8.Subject2:
                            Common.SetSpecifiedPropertyValueByName(obj, "Subject", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType8.Body:
                            Common.SetSpecifiedPropertyValueByName(obj, "Body", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType8.LastModifiedDate1:
                            string lastModifiedDateString = applicationData.Items[itemIndex].ToString();
                            if (!string.IsNullOrEmpty(lastModifiedDateString))
                            {
                                Common.SetSpecifiedPropertyValueByName(obj, "IsLastModifiedDateSpecified", true);
                                Common.SetSpecifiedPropertyValueByName(obj, "LastModifiedDate", Common.GetNoSeparatorDateTime(lastModifiedDateString));
                                Common.SetSpecifiedPropertyValueByName(obj, "LastModifiedDateString", lastModifiedDateString);
                            }

                            break;
                        case Response.ItemsChoiceType8.MessageClass1:
                            Common.SetSpecifiedPropertyValueByName(obj, "MessageClass", applicationData.Items[itemIndex]);
                            break;
                        default:
                            Common.SetSpecifiedPropertyValueByName(obj, applicationData.ItemsElementName[itemIndex].ToString(), applicationData.Items[itemIndex]);
                            break;
                    }
                }
            }

            return obj;
        }

        /// <summary>
        /// Get note instance from the ApplicationData element of Sync change response
        /// </summary>
        /// <typeparam name="T">The generic type parameter</typeparam>
        /// <param name="applicationData">The data which contains information for note</param>
        /// <returns>The returned instance</returns>
        public static T DeserializeFromChangeApplicationData<T>(Response.SyncCollectionsCollectionCommandsChangeApplicationData applicationData)
        {
            T obj = Activator.CreateInstance<T>();

            if (applicationData.ItemsElementName.Length > 0)
            {
                for (int itemIndex = 0; itemIndex < applicationData.ItemsElementName.Length; itemIndex++)
                {
                    switch (applicationData.ItemsElementName[itemIndex])
                    {
                        case Response.ItemsChoiceType7.Categories:
                        case Response.ItemsChoiceType7.Categories1:
                        case Response.ItemsChoiceType7.Categories2:
                        case Response.ItemsChoiceType7.Categories4:
                        case Response.ItemsChoiceType7.Subject:
                        case Response.ItemsChoiceType7.Subject1:
                        case Response.ItemsChoiceType7.Subject3:
                        case Response.ItemsChoiceType7.LastModifiedDate:
                        case Response.ItemsChoiceType7.MessageClass:
                            break;
                        case Response.ItemsChoiceType7.Categories3:
                            Common.SetSpecifiedPropertyValueByName(obj, "Categories", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType7.Subject2:
                            Common.SetSpecifiedPropertyValueByName(obj, "Subject", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType7.Body:
                            Common.SetSpecifiedPropertyValueByName(obj, "Body", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType7.LastModifiedDate1:
                            string lastModifiedDateString = applicationData.Items[itemIndex].ToString();
                            if (!string.IsNullOrEmpty(lastModifiedDateString))
                            {
                                Common.SetSpecifiedPropertyValueByName(obj, "IsLastModifiedDateSpecified", true);
                                Common.SetSpecifiedPropertyValueByName(obj, "LastModifiedDate", Common.GetNoSeparatorDateTime(lastModifiedDateString));
                                Common.SetSpecifiedPropertyValueByName(obj, "LastModifiedDateString", lastModifiedDateString);
                            }

                            break;
                        case Response.ItemsChoiceType7.MessageClass1:
                            Common.SetSpecifiedPropertyValueByName(obj, "MessageClass", applicationData.Items[itemIndex]);
                            break;
                        default:
                            Common.SetSpecifiedPropertyValueByName(obj, applicationData.ItemsElementName[itemIndex].ToString(), applicationData.Items[itemIndex]);
                            break;
                    }
                }
            }

            return obj;
        }
    }
}