//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------
namespace Microsoft.Protocols.TestSuites.MS_WSSREST
{
    using System.Collections.Generic;
    using System.Net;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The implement of IMS_WSSRESTSUTControlAdapter interface.
    /// </summary>
    public class MS_WSSRESTSUTControlAdapter : ManagedAdapterBase, IMS_WSSRESTSUTControlAdapter
    {
        /// <summary>
        /// The instance of MS-LISTSWS proxy class.
        /// </summary>
        private ListsSoap listsService;

        /// <summary>
        /// The list definitions.
        /// </summary>
        private List<GetListResponseGetListResult> listDefinations;

        /// <summary>
        /// Initialize SUT control adapter.
        /// </summary>
        /// <param name="testSite">The test site instance associated with the current adapter.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);

            this.listsService = new ListsSoap();

            string domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            string userName = Common.GetConfigurationPropertyValue("UserName", testSite);
            string userPassword = Common.GetConfigurationPropertyValue("Password", testSite);
            this.listsService.Credentials = new NetworkCredential(userName, userPassword, domain);
            this.listsService.Url = Common.GetConfigurationPropertyValue("ListwsServiceUrl", testSite);
            this.listsService.SoapVersion = Common.GetConfigurationPropertyValue<SoapProtocolVersion>("SoapVersion", this.Site);
            TransportProtocol transport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);

            if (transport == TransportProtocol.HTTPS)
            {
                Common.AcceptServerCertificate();
            }
        }

        /// <summary>
        /// Get the document library content type id.
        /// </summary>
        /// <param name="documentListName">The document library name.</param>
        /// <returns>The document library content type id.</returns>
        [MethodHelp("Get the document library content type id.\r\n")]
        public string GetDocumentLibraryContentTypeId(string documentListName)
        {
            GetListContentTypesResponseGetListContentTypesResult result = this.listsService.GetListContentTypes(documentListName, string.Empty);
            return result.ContentTypes.ContentTypeOrder;
        }

        /// <summary>
        /// Check whether the type of the specified field equals the expect field type.
        /// </summary>
        /// <param name="fieldName">The specified field name.</param>
        /// <param name="expectFieldType">The expect field type.</param>
        /// <returns>True if the type of the specified field name equals the expect field type, otherwise false.</returns>
        [MethodHelp("Check whether the type of the specified field equals the expect field type.\r\n")]
        public bool CheckFieldType(string fieldName, string expectFieldType)
        {
            bool result = false;
            if (null == this.listDefinations)
            {
                this.listDefinations = new List<GetListResponseGetListResult>();
                this.listDefinations.Add(this.listsService.GetList(Common.GetConfigurationPropertyValue("CalendarListName", this.Site)));
                this.listDefinations.Add(this.listsService.GetList(Common.GetConfigurationPropertyValue("DiscussionBoardListName", this.Site)));
                this.listDefinations.Add(this.listsService.GetList(Common.GetConfigurationPropertyValue("DoucmentLibraryListName", this.Site)));
                this.listDefinations.Add(this.listsService.GetList(Common.GetConfigurationPropertyValue("GeneralListName", this.Site)));
                this.listDefinations.Add(this.listsService.GetList(Common.GetConfigurationPropertyValue("SurveyListName", this.Site)));
                this.listDefinations.Add(this.listsService.GetList(Common.GetConfigurationPropertyValue("TaskListName", this.Site)));
                this.listDefinations.Add(this.listsService.GetList(Common.GetConfigurationPropertyValue("WorkflowHistoryListName", this.Site)));
            }

            foreach (GetListResponseGetListResult item in this.listDefinations)
            {
                foreach (FieldDefinition itemField in item.List.Fields.Field)
                {
                    string tempFieldName = itemField.DisplayName.Replace(" ", string.Empty).ToLower();

                    if (tempFieldName.Equals(fieldName, System.StringComparison.OrdinalIgnoreCase) && itemField.Type.Equals(expectFieldType, System.StringComparison.OrdinalIgnoreCase))
                    {
                        result = true;
                        break;
                    }
                }

                if (result == true)
                {
                    break;
                }
            }

            return result;
        }
    }
}
