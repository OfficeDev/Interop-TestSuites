namespace Microsoft.Protocols.TestSuites.Common.DataStructures
{
    using System;

    /// <summary>
    /// This class contains all the properties which are associated with the elements defined in the MS-ASCNTC protocol.
    /// </summary>
    public class Contact
    {
        #region Elements in Contacts namespace
        /// <summary>
        /// Gets or sets alias for the contact
        /// </summary>
        public string Alias { get; set; }

        /// <summary>
        /// Gets or sets anniversary date for the contact
        /// </summary>
        public DateTime? Anniversary { get; set; }

        /// <summary>
        /// Gets or sets assistant name for the contact
        /// </summary>
        public string AssistantName { get; set; }

        /// <summary>
        /// Gets or sets assistant phone number for the contact
        /// </summary>
        public string AssistantPhoneNumber { get; set; }

        /// <summary>
        /// Gets or sets birth date for the contact
        /// </summary>
        public DateTime? Birthday { get; set; }

        /// <summary>
        /// Gets or sets the business city of the contact
        /// </summary>
        public string BusinessAddressCity { get; set; }

        /// <summary>
        /// Gets or sets the business country/region of the contact
        /// </summary>
        public string BusinessAddressCountry { get; set; }

        /// <summary>
        /// Gets or sets the business postal code for the contact
        /// </summary>
        public string BusinessAddressPostalCode { get; set; }

        /// <summary>
        /// Gets or sets the business state for the contact
        /// </summary>
        public string BusinessAddressState { get; set; }

        /// <summary>
        /// Gets or sets the business street address for the contact
        /// </summary>
        public string BusinessAddressStreet { get; set; }

        /// <summary>
        /// Gets or sets the business fax number for the contact
        /// </summary>
        public string BusinessFaxNumber { get; set; }

        /// <summary>
        /// Gets or sets the primary business phone number for the contact
        /// </summary>
        public string BusinessPhoneNumber { get; set; }

        /// <summary>
        /// Gets or sets the secondary business telephone number for the contact
        /// </summary>
        public string Business2PhoneNumber { get; set; }

        /// <summary>
        /// Gets or sets the car telephone number for the contact
        /// </summary>
        public string CarPhoneNumber { get; set; }

        /// <summary>
        /// Gets or sets user labels assigned to the contact
        /// </summary>
        public Response.Categories1 Categories { get; set; }

        /// <summary>
        /// Gets or sets contact's children
        /// </summary>
        public Response.Children Children { get; set; }

        /// <summary>
        /// Gets or sets the company name for the contact
        /// </summary>
        public string CompanyName { get; set; }

        /// <summary>
        /// Gets or sets the department name for the contact
        /// </summary>
        public string Department { get; set; }

        /// <summary>
        /// Gets or sets the first e-mail address for the contact
        /// </summary>
        public string Email1Address { get; set; }

        /// <summary>
        /// Gets or sets the second e-mail address for the contact
        /// </summary>
        public string Email2Address { get; set; }

        /// <summary>
        /// Gets or sets the third e-mail address for the contact
        /// </summary>
        public string Email3Address { get; set; }

        /// <summary>
        /// Gets or sets FileAs for the contact
        /// </summary>
        public string FileAs { get; set; }

        /// <summary>
        /// Gets or sets the first name of the contact
        /// </summary>
        public string FirstName { get; set; }

        /// <summary>
        /// Gets or sets the home city for the contact
        /// </summary>
        public string HomeAddressCity { get; set; }

        /// <summary>
        /// Gets or sets the home country/region for the contact
        /// </summary>
        public string HomeAddressCountry { get; set; }

        /// <summary>
        /// Gets or sets the home postal code for the contact
        /// </summary>
        public string HomeAddressPostalCode { get; set; }

        /// <summary>
        /// Gets or sets the home state for the contact
        /// </summary>
        public string HomeAddressState { get; set; }

        /// <summary>
        /// Gets or sets the home street address for the contact
        /// </summary>
        public string HomeAddressStreet { get; set; }

        /// <summary>
        /// Gets or sets the home fax number for the contact
        /// </summary>
        public string HomeFaxNumber { get; set; }

        /// <summary>
        /// Gets or sets the home phone number for the contact
        /// </summary>
        public string HomePhoneNumber { get; set; }

        /// <summary>
        /// Gets or sets the alternative home phone number for the contact
        /// </summary>
        public string Home2PhoneNumber { get; set; }

        /// <summary>
        /// Gets or sets the contact's job title
        /// </summary>
        public string JobTitle { get; set; }

        /// <summary>
        /// Gets or sets the contact's last name
        /// </summary>
        public string LastName { get; set; }

        /// <summary>
        /// Gets or sets the middle name of the contact
        /// </summary>
        public string MiddleName { get; set; }

        /// <summary>
        /// Gets or sets the mobile phone number for the contact
        /// </summary>
        public string MobilePhoneNumber { get; set; }

        /// <summary>
        /// Gets or sets the office location for the contact
        /// </summary>
        public string OfficeLocation { get; set; }

        /// <summary>
        /// Gets or sets the city for the contact's alternate address
        /// </summary>
        public string OtherAddressCity { get; set; }

        /// <summary>
        /// Gets or sets the country/region of the contact's alternate address
        /// </summary>
        public string OtherAddressCountry { get; set; }

        /// <summary>
        /// Gets or sets the postal code of the contact's alternate address
        /// </summary>
        public string OtherAddressPostalCode { get; set; }

        /// <summary>
        /// Gets or sets the state of the contact's alternate address
        /// </summary>
        public string OtherAddressState { get; set; }

        /// <summary>
        /// Gets or sets the street address of the contact's alternate address
        /// </summary>
        public string OtherAddressStreet { get; set; }

        /// <summary>
        /// Gets or sets the pager number for the contact
        /// </summary>
        public string PagerNumber { get; set; }

        /// <summary>
        /// Gets or sets the picture of the contact
        /// </summary>
        public string Picture { get; set; }

        /// <summary>
        /// Gets or sets the radio phone number for the contact
        /// </summary>
        public string RadioPhoneNumber { get; set; }

        /// <summary>
        /// Gets or sets the contact's spouse/partner
        /// </summary>
        public string Spouse { get; set; }

        /// <summary>
        /// Gets or sets the suffix for the contact's name
        /// </summary>
        public string Suffix { get; set; }

        /// <summary>
        /// Gets or sets the contact's business title
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets the Web site or personal Web page for the contact
        /// </summary>
        public string WebPage { get; set; }

        /// <summary>
        /// Gets or sets the rank of this contact entry in the recipient information cache
        /// </summary>
        public int? WeightedRank { get; set; }

        /// <summary>
        /// Gets or sets the Japanese phonetic rendering of the company name for the contact
        /// </summary>
        public string YomiCompanyName { get; set; }

        /// <summary>
        /// Gets or sets the Japanese phonetic rendering of the first name of the contact
        /// </summary>
        public string YomiFirstName { get; set; }

        /// <summary>
        /// Gets or sets the Japanese phonetic rendering of the last name of the contact
        /// </summary>
        public string YomiLastName { get; set; }
        #endregion

        #region Elements in Contacts2 namespace
        /// <summary>
        /// Gets or sets account name for the contact
        /// </summary>
        public string AccountName { get; set; }

        /// <summary>
        /// Gets or sets the main telephone number for the contact's company
        /// </summary>
        public string CompanyMainPhone { get; set; }

        /// <summary>
        /// Gets or sets the customer identifier (ID) for the contact
        /// </summary>
        public string CustomerId { get; set; }

        /// <summary>
        /// Gets or sets the government-assigned identifier (ID) for the contact
        /// </summary>
        public string GovernmentId { get; set; }

        /// <summary>
        /// Gets or sets the instant messaging address for the contact
        /// </summary>
        public string IMAddress { get; set; }

        /// <summary>
        /// Gets or sets the alternative instant messaging address for the contact
        /// </summary>
        public string IMAddress2 { get; set; }

        /// <summary>
        /// Gets or sets the tertiary instant messaging address for the contact
        /// </summary>
        public string IMAddress3 { get; set; }

        /// <summary>
        /// Gets or sets the distinguished name (DN) (1) of the contact's manager
        /// </summary>
        public string ManagerName { get; set; }

        /// <summary>
        /// Gets or sets the Multimedia Messaging Service (MMS) address for the contact
        /// </summary>
        public string MMS { get; set; }

        /// <summary>
        /// Gets or sets the nickname for the contact
        /// </summary>
        public string NickName { get; set; }
        #endregion

        #region Elements in AirSyncBase namespace
        /// <summary>
        /// Gets or sets notes for the contact
        /// </summary>
        public Response.Body Body { get; set; }
        #endregion

        /// <summary>
        /// Deserialize to object instance from SyncCollectionsCollectionCommandsAddApplicationData.
        /// </summary>
        /// <typeparam name="T">The generic type parameter.</typeparam>
        /// <param name="applicationData">The application data which contains new added information.</param>
        /// <returns>The returned instance.</returns>
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
                        case Response.ItemsChoiceType8.Categories2:
                        case Response.ItemsChoiceType8.Categories3:
                        case Response.ItemsChoiceType8.Categories4:
                            break;
                        case Response.ItemsChoiceType8.Categories1:
                            Common.SetSpecifiedPropertyValueByName(obj, "Categories", applicationData.Items[itemIndex]);
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
        /// Deserialize to object instance from SyncCollectionsCollectionCommandsChangeApplicationData.
        /// </summary>
        /// <typeparam name="T">The generic type parameter.</typeparam>
        /// <param name="applicationData">The application data which contains changes information.</param>
        /// <returns>The returned instance.</returns>
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
                        case Response.ItemsChoiceType7.Categories2:
                        case Response.ItemsChoiceType7.Categories3:
                        case Response.ItemsChoiceType7.Categories4:
                            break;

                        case Response.ItemsChoiceType7.Categories1:
                            Common.SetSpecifiedPropertyValueByName(obj, "Categories", applicationData.Items[itemIndex]);
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
        /// Deserialize to object instance from ItemOperations Properties.
        /// </summary>
        /// <typeparam name="T">The generic type parameter.</typeparam>
        /// <param name="fetchProperties">The Properties data which contains new added information.</param>
        /// <returns>The object instance.</returns>
        public static T DeserializeFromFetchProperties<T>(Response.Properties fetchProperties)
        {
            T obj = Activator.CreateInstance<T>();
            if (fetchProperties.ItemsElementName.Length > 0)
            {
                for (int itemIndex = 0; itemIndex < fetchProperties.ItemsElementName.Length; itemIndex++)
                {
                    switch (fetchProperties.ItemsElementName[itemIndex])
                    {
                        case Response.ItemsChoiceType3.Categories:
                        case Response.ItemsChoiceType3.Categories2:
                        //case Response.ItemsChoiceType3.Categories3:
                        case Response.ItemsChoiceType3.Categories4:
                            break;
                        case Response.ItemsChoiceType3.Categories1:
                            Common.SetSpecifiedPropertyValueByName(obj, "Categories", fetchProperties.Items[itemIndex]);
                            break;
                        default:
                            Common.SetSpecifiedPropertyValueByName(obj, fetchProperties.ItemsElementName[itemIndex].ToString(), fetchProperties.Items[itemIndex]);
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
                        case Response.ItemsChoiceType6.Categories1:
                        case Response.ItemsChoiceType6.Categories2:
                        case Response.ItemsChoiceType6.Categories3:
                        case Response.ItemsChoiceType6.FirstName1:
                        case Response.ItemsChoiceType6.LastName1:
                        case Response.ItemsChoiceType6.Title1:
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