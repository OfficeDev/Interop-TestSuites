namespace Microsoft.Protocols.TestSuites.Common.DataStructures
{
    using System;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This class contains all properties which are associated with the elements defined in the MS-ASTASK protocol.
    /// </summary>
    public class Task
    {
        /// <summary>
        /// Gets or sets Task Body information of the Task
        /// </summary>
        public Response.Body Body { get; set; }

        /// <summary>
        /// Gets or sets Subject information of the Task
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// Gets or sets Importance information of the Task
        /// </summary>
        public byte? Importance { get; set; }

        /// <summary>
        /// Gets or sets UtcStartDate information of the Task
        /// </summary>
        public DateTime? UtcStartDate { get; set; }

        /// <summary>
        /// Gets or sets StartDate information of the Task
        /// </summary>
        public DateTime? StartDate { get; set; }

        /// <summary>
        /// Gets or sets UtcDueDate information of the Task
        /// </summary>
        public DateTime? UtcDueDate { get; set; }

        /// <summary>
        /// Gets or sets DueDate information of the Task
        /// </summary>
        public DateTime? DueDate { get; set; }

        /// <summary>
        /// Gets or sets Categories information of the Task
        /// </summary>
        public Response.Categories3 Categories { get; set; }

        /// <summary>
        /// Gets or sets Recurrence information of the Task
        /// </summary>
        public Response.Recurrence1 Recurrence { get; set; }

        /// <summary>
        /// Gets or sets Complete information of the Task
        /// </summary>
        public byte? Complete { get; set; }

        /// <summary>
        /// Gets or sets DateCompleted information of the Task
        /// </summary>
        public DateTime? DateCompleted { get; set; }

        /// <summary>
        /// Gets or sets Sensitivity information of the Task
        /// </summary>
        public byte? Sensitivity { get; set; }

        /// <summary>
        /// Gets or sets ReminderTime information of the Task
        /// </summary>
        public DateTime? ReminderTime { get; set; }

        /// <summary>
        /// Gets or sets ReminderSet information of the Task
        /// </summary>
        public byte? ReminderSet { get; set; }

        /// <summary>
        /// Desterilize to object instance from Properties
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
                            break;
                        case Response.ItemsChoiceType3.Categories1:
                            break;
                        case Response.ItemsChoiceType3.Categories2:
                            break;
                        //case Response.ItemsChoiceType3.Categories3:
                        //    break;
                        case Response.ItemsChoiceType3.Subject:
                            break;
                        case Response.ItemsChoiceType3.Subject1:
                            break;
                        //case Response.ItemsChoiceType3.Subject2:
                        //    break;
                        case Response.ItemsChoiceType3.Recurrence:
                            break;
                        case Response.ItemsChoiceType3.Categories4:
                            Common.SetSpecifiedPropertyValueByName(obj, "Categories", properties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType3.Body:
                            Common.SetSpecifiedPropertyValueByName(obj, "Body", properties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType3.Importance1:
                            Common.SetSpecifiedPropertyValueByName(obj, "Importance", properties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType3.Sensitivity1:
                            Common.SetSpecifiedPropertyValueByName(obj, "Sensitivity", properties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType3.Subject3:
                            Common.SetSpecifiedPropertyValueByName(obj, "Subject", properties.Items[itemIndex]);
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
        /// Get task instance from the Properties element of Search response
        /// </summary>
        /// <typeparam name="T">The generic type parameter</typeparam>
        /// <param name="properties">The data which contains information for task</param>
        /// <returns>The returned task instance</returns>
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
                            break;
                        case Response.ItemsChoiceType6.Categories1:
                            break;
                        case Response.ItemsChoiceType6.Categories2:
                            break;
                        case Response.ItemsChoiceType6.Subject:
                            break;
                        case Response.ItemsChoiceType6.Subject1:
                            break;
                        case Response.ItemsChoiceType6.Subject2:
                            break;
                        case Response.ItemsChoiceType6.MessageClass:
                            break;
                        case Response.ItemsChoiceType6.Recurrence:
                            break;
                        case Response.ItemsChoiceType6.Categories3:
                            Common.SetSpecifiedPropertyValueByName(obj, "Categories", properties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType6.Importance1:
                            Common.SetSpecifiedPropertyValueByName(obj, "Importance", properties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType6.Sensitivity1:
                            Common.SetSpecifiedPropertyValueByName(obj, "Sensitivity", properties.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType6.Subject3:
                            Common.SetSpecifiedPropertyValueByName(obj, "Subject", properties.Items[itemIndex]);
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
        /// Get task instance from the ApplicationData element of Sync add response
        /// </summary>
        /// <typeparam name="T">The generic type parameter</typeparam>
        /// <param name="applicationData">The data which contains information for task</param>
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
                            break;
                        case Response.ItemsChoiceType8.Categories1:
                            break;
                        case Response.ItemsChoiceType8.Categories2:
                            break;
                        case Response.ItemsChoiceType8.Categories3:
                            break;
                        case Response.ItemsChoiceType8.Subject:
                            break;
                        case Response.ItemsChoiceType8.Subject1:
                            break;
                        case Response.ItemsChoiceType8.Subject2:
                            break;
                        case Response.ItemsChoiceType8.Recurrence:
                            break;
                        case Response.ItemsChoiceType8.Recurrence1:
                            Common.SetSpecifiedPropertyValueByName(obj, "Recurrence", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType8.Categories4:
                            Common.SetSpecifiedPropertyValueByName(obj, "Categories", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType8.Importance1:
                            Common.SetSpecifiedPropertyValueByName(obj, "Importance", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType8.Sensitivity1:
                            Common.SetSpecifiedPropertyValueByName(obj, "Sensitivity", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType8.Subject3:
                            Common.SetSpecifiedPropertyValueByName(obj, "Subject", applicationData.Items[itemIndex]);
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
        /// Get task instance from the ApplicationData element of Sync change response
        /// </summary>
        /// <typeparam name="T">The generic type parameter</typeparam>
        /// <param name="applicationData">The data which contains information for task</param>
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
                            break;
                        case Response.ItemsChoiceType7.Categories1:
                            break;
                        case Response.ItemsChoiceType7.Categories2:
                            break;
                        case Response.ItemsChoiceType7.Categories3:
                            break;
                        case Response.ItemsChoiceType7.Subject:
                            break;
                        case Response.ItemsChoiceType7.Subject1:
                            break;
                        case Response.ItemsChoiceType7.Subject2:
                            break;
                        case Response.ItemsChoiceType7.Recurrence:
                            break;
                        case Response.ItemsChoiceType7.Categories4:
                            Common.SetSpecifiedPropertyValueByName(obj, "Categories", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType7.Importance1:
                            Common.SetSpecifiedPropertyValueByName(obj, "Importance", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType7.Sensitivity1:
                            Common.SetSpecifiedPropertyValueByName(obj, "Sensitivity", applicationData.Items[itemIndex]);
                            break;
                        case Response.ItemsChoiceType7.Subject3:
                            Common.SetSpecifiedPropertyValueByName(obj, "Subject", applicationData.Items[itemIndex]);
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