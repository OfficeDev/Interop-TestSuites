//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSATT
{
    using System.Collections.Generic;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods defined in MS-OXWSATT.
    /// </summary>
    public interface IMS_OXWSATTAdapter : IAdapter
    {
        /// <summary>
        /// Gets the raw XML request sent to protocol SUT.
        /// </summary>
        IXPathNavigable LastRawRequestXml { get; }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT.
        /// </summary>
        IXPathNavigable LastRawResponseXml { get; }

        /// <summary>
        /// Creates an item or file attachment on an item in the server store. 
        /// </summary>
        /// <param name="createAttachmentRequest">A CreateAttachmentType complex type specifies a request message to attach an item or file to a specified item in the server database. </param>
        /// <returns>A CreateAttachmentResponseType complex type specifies the response message that is returned by the CreateAttachment operation. </returns>
        CreateAttachmentResponseType CreateAttachment(CreateAttachmentType createAttachmentRequest);

        /// <summary>
        /// Gets an attachment from an item in the server store.
        /// </summary>
        /// <param name="getAttachmentRequest">A GetAttachmentType complex type specifies a request message to get attached items and files on an item in the server database.</param>
        /// <returns>A GetAttachmentResponseType complex type specifies the response message that is returned by the GetAttachment operation.</returns>
        GetAttachmentResponseType GetAttachment(GetAttachmentType getAttachmentRequest);

        /// <summary>
        /// Deletes an attachment from an item in the server store. 
        /// </summary>
        /// <param name="deleteAttachmentRequest">A DeleteAttachmentType complex type specifies a request message to delete an attachment on an item in the server database.</param>
        /// <returns>A DeleteAttachmentResponseType complex type specifies the response message that is returned by the DeleteAttachment operation.</returns>
        DeleteAttachmentResponseType DeleteAttachment(DeleteAttachmentType deleteAttachmentRequest);

        /// <summary>
        /// Configure the SOAP header before calling operations.
        /// </summary>
        /// <param name="headerValues">Specify the values for SOAP header.</param>
        void ConfigureSOAPHeader(Dictionary<string, object> headerValues);
    }
}