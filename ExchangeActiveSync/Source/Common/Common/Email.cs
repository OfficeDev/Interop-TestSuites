namespace Microsoft.Protocols.TestSuites.Common.DataStructures
{
    using System;

    /// <summary>
    /// The class contains all email properties
    /// </summary>
    public class Email
    {
        #region Elements from Email namespace
        /// <summary>
        /// Gets or sets To
        /// </summary>
        public string To { get; set; }

        /// <summary>
        /// Gets or sets From
        /// </summary>
        public string From { get; set; }

        /// <summary>
        /// Gets or sets CC
        /// </summary>
        public string CC { get; set; }

        /// <summary>
        /// Gets or sets Subject
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// Gets or sets ReplyTo
        /// </summary>
        public string ReplyTo { get; set; }

        /// <summary>
        /// Gets or sets DateReceived
        /// </summary>
        public DateTime? DateReceived { get; set; }

        /// <summary>
        /// Gets or sets DisplayTo
        /// </summary>
        public string DisplayTo { get; set; }

        /// <summary>
        /// Gets or sets Importance
        /// </summary>
        public byte? Importance { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether email is read
        /// </summary>
        public bool? Read { get; set; }

        /// <summary>
        /// Gets or sets MessageClass
        /// </summary>
        public string MessageClass { get; set; }

        /// <summary>
        /// Gets or sets MeetingRequest
        /// </summary>
        public Response.MeetingRequest MeetingRequest { get; set; }

        /// <summary>
        /// Gets or sets ThreadTopic
        /// </summary>
        public string ThreadTopic { get; set; }

        /// <summary>
        /// Gets or sets InternetCPID
        /// </summary>
        public string InternetCPID { get; set; }

        /// <summary>
        /// Gets or sets Flag
        /// </summary>
        public Response.Flag Flag { get; set; }

        /// <summary>
        /// Gets or sets ContentClass
        /// </summary>
        public string ContentClass { get; set; }

        /// <summary>
        /// Gets or sets Categories
        /// </summary>
        public Response.Categories2 Categories { get; set; }
        #endregion

        #region Elements from AirSyncBase namespace
        /// <summary>
        /// Gets or sets Body
        /// </summary>
        public Response.Body Body { get; set; }

        /// <summary>
        /// Gets or sets BodyPart
        /// </summary>
        public Response.BodyPart BodyPart { get; set; }

        /// <summary>
        /// Gets or sets Attachments
        /// </summary>
        public Response.Attachments Attachments { get; set; }

        /// <summary>
        /// Gets or sets NativeBodyType
        /// </summary>
        public byte? NativeBodyType { get; set; }
        #endregion

        #region Elements from Email2 namespace
        /// <summary>
        /// Gets or sets BCC
        /// </summary>
        public string Bcc { get; set; }
        /// <summary>
        /// Gets or sets IsDraft
        /// </summary>
        public bool? IsDraft { get; set; } 
        /// <summary>
        /// Gets or sets UmCallerID
        /// </summary>
        public string UmCallerID { get; set; }

        /// <summary>
        /// Gets or sets UmUserNotes
        /// </summary>
        public string UmUserNotes { get; set; }

        /// <summary>
        /// Gets or sets ConversationId
        /// </summary>
        public string ConversationId { get; set; }

        /// <summary>
        /// Gets or sets ConversationIndex
        /// </summary>
        public string ConversationIndex { get; set; }

        /// <summary>
        /// Gets or sets LastVerbExecuted
        /// </summary>
        public int? LastVerbExecuted { get; set; }

        /// <summary>
        /// Gets or sets LastVerbExecutionTime
        /// </summary>
        public DateTime? LastVerbExecutionTime { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether ReceivedAsBcc element include in response
        /// </summary>
        public bool? ReceivedAsBcc { get; set; }

        /// <summary>
        /// Gets or sets Sender
        /// </summary>
        public string Sender { get; set; }

        /// <summary>
        /// Gets or sets AccountId
        /// </summary>
        public string AccountId { get; set; }

        #endregion

        #region Element from RightsManagement
        /// <summary>
        /// Gets or sets RightsManagementLicense
        /// </summary>
        public Response.RightsManagementLicense RightsManagementLicense { get; set; }
        #endregion

        /// <summary>
        /// Gets or sets a value indicating whether a Read tag includes in response
        /// </summary>
        public bool ReadIsInclude { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether a Flag tag includes in response
        /// </summary>
        public bool FlagIsInclude { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether a Categories tag includes in response
        /// </summary>
        public bool CategoriesIsInclude { get; set; }

        /// <summary>
        /// Deserialize to Email instance from Response.Properties.
        /// </summary>
        /// <typeparam name="T">The generic type parameter</typeparam>
        /// <param name="properties">The application data which contains Properties element from the ItemOperations command response.</param>
        /// <returns>The returned instance.</returns>
        public static T DeserializeFromFetchProperties<T>(Response.Properties properties)
        {
            T obj = Activator.CreateInstance<T>();

            if (properties != null && properties.ItemsElementName.Length > 0)
            {
                for (int itemIndex = 0; itemIndex < properties.ItemsElementName.Length; itemIndex++)
                {
                    switch (properties.ItemsElementName[itemIndex])
                    {
                        case Response.ItemsChoiceType3.Categories:
                        case Response.ItemsChoiceType3.Categories1:
                        //case Response.ItemsChoiceType3.Categories3:
                        case Response.ItemsChoiceType3.Categories4:
                        case Response.ItemsChoiceType3.Importance1:
                        //case Response.ItemsChoiceType3.MessageClass1:
                        case Response.ItemsChoiceType3.Subject:
                        //case Response.ItemsChoiceType3.Subject2:
                        case Response.ItemsChoiceType3.Subject3:
                            break;
                        case Response.ItemsChoiceType3.Read:
                            Common.SetSpecifiedPropertyValueByName(obj, "ReadIsInclude", true);
                            Common.SetSpecifiedPropertyValueByName(obj, "Read", properties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType3.Flag:
                            Common.SetSpecifiedPropertyValueByName(obj, "FlagIsInclude", true);
                            Common.SetSpecifiedPropertyValueByName(obj, "Flag", properties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType3.Categories2:
                            Common.SetSpecifiedPropertyValueByName(obj, "Categories", properties.Items[itemIndex]);
                            Common.SetSpecifiedPropertyValueByName(obj, "CategoriesIsInclude", true);
                            break;
                        case Response.ItemsChoiceType3.Subject1:
                            Common.SetSpecifiedPropertyValueByName(obj, "Subject", properties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType3.Bcc:
                            Common.SetSpecifiedPropertyValueByName(obj, "Bcc", properties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType3.IsDraft:
                            Common.SetSpecifiedPropertyValueByName(obj, "IsDraft", properties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType3.Cc:
                            Common.SetSpecifiedPropertyValueByName(obj, "CC", properties.Items[itemIndex]);
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
        /// Deserialize to object instance from SyncCollectionsCollectionCommandsAddApplicationData"
        /// </summary>
        /// <typeparam name="T">The generic type parameter</typeparam>
        /// <param name="applicationData">The application data which contains new added information</param>
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
                        case Response.ItemsChoiceType8.Categories3:
                        case Response.ItemsChoiceType8.Categories4:
                        case Response.ItemsChoiceType8.Importance1:
                        case Response.ItemsChoiceType8.MessageClass1:
                        case Response.ItemsChoiceType8.Subject:
                        case Response.ItemsChoiceType8.Subject2:
                        case Response.ItemsChoiceType8.Subject3:
                            break;
                        case Response.ItemsChoiceType8.Read:
                            Common.SetSpecifiedPropertyValueByName(obj, "ReadIsInclude", true);
                            Common.SetSpecifiedPropertyValueByName(obj, "Read", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType8.Flag:
                            Common.SetSpecifiedPropertyValueByName(obj, "FlagIsInclude", true);
                            Common.SetSpecifiedPropertyValueByName(obj, "Flag", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType8.Categories2:
                            Common.SetSpecifiedPropertyValueByName(obj, "Categories", applicationData.Items[itemIndex]);
                            Common.SetSpecifiedPropertyValueByName(obj, "CategoriesIsInclude", true);
                            break;
                        case Response.ItemsChoiceType8.Subject1:
                            Common.SetSpecifiedPropertyValueByName(obj, "Subject", applicationData.Items[itemIndex]);
                            break; 
                        case Response.ItemsChoiceType8.Bcc:
                            Common.SetSpecifiedPropertyValueByName(obj, "Bcc", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType8.IsDraft:
                            Common.SetSpecifiedPropertyValueByName(obj, "IsDraft", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType8.Cc:
                            Common.SetSpecifiedPropertyValueByName(obj, "CC", applicationData.Items[itemIndex]);
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
        /// Deserialize to object instance from SyncCollectionsCollectionCommandsChangeApplicationData
        /// </summary>
        /// <typeparam name="T">The generic type parameter</typeparam>
        /// <param name="applicationData">The application data which contains changes information</param>
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
                        case Response.ItemsChoiceType7.Categories3:
                        case Response.ItemsChoiceType7.Categories4:
                        case Response.ItemsChoiceType7.Importance1:
                        case Response.ItemsChoiceType7.MessageClass1:
                        case Response.ItemsChoiceType7.Subject:
                        case Response.ItemsChoiceType7.Subject2:
                        case Response.ItemsChoiceType7.Subject3:
                            break;
                        case Response.ItemsChoiceType7.Read:
                            Common.SetSpecifiedPropertyValueByName(obj, "ReadIsInclude", true);
                            Common.SetSpecifiedPropertyValueByName(obj, "Read", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType7.Flag:
                            Common.SetSpecifiedPropertyValueByName(obj, "FlagIsInclude", true);
                            Common.SetSpecifiedPropertyValueByName(obj, "Flag", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType7.Categories2:
                            Common.SetSpecifiedPropertyValueByName(obj, "Categories", applicationData.Items[itemIndex]);
                            Common.SetSpecifiedPropertyValueByName(obj, "CategoriesIsInclude", true);
                            break;
                        case Response.ItemsChoiceType7.Subject1:
                            Common.SetSpecifiedPropertyValueByName(obj, "Subject", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType7.Bcc:
                            Common.SetSpecifiedPropertyValueByName(obj, "Bcc", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType7.IsDraft:
                            Common.SetSpecifiedPropertyValueByName(obj, "IsDraft", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType7.Cc:
                            Common.SetSpecifiedPropertyValueByName(obj, "CC", applicationData.Items[itemIndex]);
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
        /// Deserialize to Email instance from Response.SearchResponseStoreResultProperties.
        /// </summary>
        /// <typeparam name="T">The generic type parameter.</typeparam>
        /// <param name="searchResultProperties">The application data which contains Properties element from the Search command response.</param>
        /// <returns>The returned instance.</returns>
        public static T DeserializeFromSearchProperties<T>(Response.SearchResponseStoreResultProperties searchResultProperties)
        {
            T obj = Activator.CreateInstance<T>();
            if (searchResultProperties != null && searchResultProperties.ItemsElementName.Length > 0)
            {
                for (int itemIndex = 0; itemIndex < searchResultProperties.ItemsElementName.Length; itemIndex++)
                {
                    switch (searchResultProperties.ItemsElementName[itemIndex])
                    {
                        case Response.ItemsChoiceType6.Categories:
                        case Response.ItemsChoiceType6.Categories2:
                        case Response.ItemsChoiceType6.Categories3:
                        case Response.ItemsChoiceType6.Importance1:
                        case Response.ItemsChoiceType6.MessageClass1:
                        case Response.ItemsChoiceType6.Subject:
                        case Response.ItemsChoiceType6.Subject2:
                        case Response.ItemsChoiceType6.Subject3:
                            break;
                        case Response.ItemsChoiceType6.Read:
                            Common.SetSpecifiedPropertyValueByName(obj, "ReadIsInclude", true);
                            Common.SetSpecifiedPropertyValueByName(obj, "Read", searchResultProperties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType6.Flag:
                            Common.SetSpecifiedPropertyValueByName(obj, "FlagIsInclude", true);
                            Common.SetSpecifiedPropertyValueByName(obj, "Flag", searchResultProperties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType6.Categories1:
                            Common.SetSpecifiedPropertyValueByName(obj, "Categories", searchResultProperties.Items[itemIndex]);
                            Common.SetSpecifiedPropertyValueByName(obj, "CategoriesIsInclude", true);
                            break;
                        case Response.ItemsChoiceType6.Subject1:
                            Common.SetSpecifiedPropertyValueByName(obj, "Subject", searchResultProperties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType6.Bcc:
                            Common.SetSpecifiedPropertyValueByName(obj, "Bcc", searchResultProperties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType6.IsDraft:
                            Common.SetSpecifiedPropertyValueByName(obj, "IsDraft", searchResultProperties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType6.Cc:
                            Common.SetSpecifiedPropertyValueByName(obj, "CC", searchResultProperties.Items[itemIndex]);
                            break;
                        default:
                            Common.SetSpecifiedPropertyValueByName(obj, searchResultProperties.ItemsElementName[itemIndex].ToString(), searchResultProperties.Items[itemIndex]);
                            break;
                    }
                }
            }

            return obj;
        }
    }
}