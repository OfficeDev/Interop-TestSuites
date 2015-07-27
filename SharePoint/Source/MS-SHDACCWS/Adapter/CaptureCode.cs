//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_SHDACCWS
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This adapter class of MS-SHDACCWS 
    /// </summary>
    public partial class MS_SHDACCWSAdapter : ManagedAdapterBase, IMS_SHDACCWSAdapter
    {
        /// <summary>
        /// The method is used to capture requirements about schema of IsOnlyClient Operation
        /// </summary>
        private void VerifySchemaOfIsOnlyClientOperation()
        {
            // Verify MS-SHDACCWS requirement: MS-SHDACCWS_R31
            // If no exception thrown, the schema of IsOnlyClient operation is valid,so capture R31 directly.
            this.Site.CaptureRequirement(
                                   31,
                                   @"[The definition of the IsOnlyClient method is as follows] 
                                   <wsdl:operation name=""IsOnlyClient"">
                                    <wsdl:input message=""tns:IsOnlyClientSoapIn"" />
                                    <wsdl:output message=""tns:IsOnlyClientSoapOut"" />
                                   </wsdl:operation>");

            // Verify MS-SHDACCWS requirement: MS-SHDACCWS_R34
            // If no exception thrown, the schema of IsOnlyClient operation is valid,so capture R34 directly.
            this.Site.CaptureRequirement(
                                    34,
                                    @"[In IsOnlyClient] The protocol server responds with an IsOnlyClientSoapOut response message.");

            // Verify MS-SHDACCWS requirement: MS-SHDACCWS_R41
            // If no exception thrown, the schema of IsOnlyClient operation is valid,so capture R41 directly.
            this.Site.CaptureRequirement(
                                     41,
                                     @"[In IsOnlyClientSoapOut] The SOAP body contains an IsOnlyClientResponse element.");

            // Verify MS-SHDACCWS requirement: MS-SHDACCWS_R50
            // If no exception thrown, the schema of IsOnlyClient operation is valid,so capture R50 directly.
            this.Site.CaptureRequirement(
                                     50,
                                        @"[The definition of the IsOnlyClientResponse element is as follows] 
<s:element name=""IsOnlyClientResponse"">
  <s:complexType>
    <s:sequence>
      <s:element minOccurs=""1"" maxOccurs=""1"" name=""IsOnlyClientResult"" type=""s:boolean"" />
    </s:sequence>
  </s:complexType>
</s:element>");
        }

        /// <summary>
        /// The method is used to capture HTTP/HTTPS SOAP requirements.
        /// </summary>
        private void VerifyTransportProtocol()
        {
            TransportProtocol currentTransport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);
            switch (currentTransport)
            {
                case TransportProtocol.HTTP:
                    {
                        // If current test suite use HTTP as the low level transport protocol, then capture R1.
                        // Verify MS-SHDACCWS requirement: MS-SHDACCWS_R1
                        this.Site.CaptureRequirement(
                                                1,
                                                @"[In Transport] Protocol servers MUST support SOAP over HTTP.");
                        break;
                    }

                case TransportProtocol.HTTPS:
                    {
                        if (Common.IsRequirementEnabled(60, this.Site))
                        {
                            // If current test suite use HTTPS as the low level transport protocol, then capture R60.
                            // Verify MS-SHDACCWS requirement: MS-SHDACCWS_R60
                            this.Site.CaptureRequirement(
                                               60,
                                               @"[Transport:] 
                                               Implementation does additionally support SOAP over HTTPS for securing communication with clients.
                                               (Microsoft SharePoint Foundation 2010 and Microsoft SharePoint Foundation 2013 products follow this behavior.)");
                        }

                        break;
                    }

                default:
                    {
                        this.Site.Assert.Fail("Transport: {0} is not HTTP or HTTPS", currentTransport);
                        break;
                    }
            }

            // Verify MS-SHDACCWS requirement: MS-SHDACCWS_R3
            SoapVersion currentSoapValue = Common.GetConfigurationPropertyValue<SoapVersion>("SOAPVersion", this.Site);
            switch (currentSoapValue)
            {
                case SoapVersion.SOAP11:
                case SoapVersion.SOAP12:
                    {
                        // If current test suite use SOAP1.1 or SOAP1.2 as the soap version, then capture R3.
                        // Verify MS-WWSP requirement: MS-WWSP_R3
                        this.Site.CaptureRequirement(
                                               3,
                                               @"[In Transport] Protocol messages MUST be formatted as specified either in [SOAP1.1], Section 4, SOAP envelope or in [SOAP1.2/1],  Section 5, SOAP Message Construct.");
                        break;
                    }

                default:
                    {
                        this.Site.Assert.Fail("Message format: {0} is not SOAP11 or SOAP12", currentSoapValue);
                        break;
                    }
            }
        }
    }
}