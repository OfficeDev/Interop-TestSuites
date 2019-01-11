namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Xml;
    using System.Xml.Serialization;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This class is used to validate the schema for each specified sub response.
    /// </summary>
    public class SubResponseValidationWrapper
    {
        /// <summary>
        /// Gets or sets the sub request token.
        /// </summary>
        public string SubToken { get; set; }

        /// <summary>
        /// Gets or sets the sub request type.
        /// </summary>
        public string SubRequestType { get; set; }

        /// <summary>
        /// Gets the sub response element name, which is corresponding with the value SubRequestType.
        /// </summary>
        public string SubResponseElementName
        {
            get
            {
                if (this.SubRequestType == null)
                {
                    return null;
                }

                return this.SubRequestType.Replace("RequestType", "Response");
            }
        }

        /// <summary>
        /// Gets the sub response type name, which is corresponding with the value SubRequestType.
        /// </summary>
        public string SubResponseTypeName
        {
            get
            {
                if (this.SubRequestType == null)
                {
                    return null;
                }

                return this.SubRequestType.Replace("Request", "Response");
            }
        }

        /// <summary>
        /// This method is used to validate the sub response according to the current record sub request token and sub request type.
        /// </summary>
        /// <param name="rawResponse">Specify the raw XML string response returned by the protocol server.</param>
        /// <param name="site">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        public void Validate(string rawResponse, ITestSite site)
        {
            // Extract the sub response whose token equals the SubToken value.
            XmlDocument subResponseDocument = this.ExtractSubResponseNode(rawResponse);

            // De-serialize the sub response to instance
            object subResponse = this.SerializeSubResponse(subResponseDocument, site);

            // Try to parse the MS-FSSHTTPB structure
            if (subResponse is CellSubResponseType)
            {
                // If the sub request type is CellSubRequestType, then indicating that there is one MS-FSSHTTPB response embedded. Try parse this an capture all the related requirements.
                CellSubResponseType cellSubResponse = subResponse as CellSubResponseType;
                if (cellSubResponse.SubResponseData != null && cellSubResponse.SubResponseData.Text.Length == 1)
                {
                    string subResponseBase64 = cellSubResponse.SubResponseData.Text[0];
                    byte[] subResponseBinary = Convert.FromBase64String(subResponseBase64);
                    FsshttpbResponse fsshttpbResponse = FsshttpbResponse.DeserializeResponseFromByteArray(subResponseBinary, 0);

                    if (fsshttpbResponse.DataElementPackage != null && fsshttpbResponse.DataElementPackage.DataElements != null)
                    {
                        // If the response data elements is complete, then try to verify the requirements related in the MS-FSSHTPD
                        foreach (DataElement storageIndex in fsshttpbResponse.DataElementPackage.DataElements.Where(dataElement => dataElement.DataElementType == DataElementType.StorageIndexDataElementData))
                        {
                            // Just build the root node to try to parse the signature related requirements, no need to restore the result.
                            new IntermediateNodeObject.RootNodeObjectBuilder().Build(
                                       fsshttpbResponse.DataElementPackage.DataElements,
                                       storageIndex.DataElementExtendedGUID);
                        }

                        if (SharedContext.Current.FileUrl.ToLowerInvariant().EndsWith(".one")
                            || SharedContext.Current.FileUrl.ToLowerInvariant().EndsWith(".onetoc2"))
                        {
                            MSONESTOREParser onenoteParser = new MSONESTOREParser();
                            MSOneStorePackage package = onenoteParser.Parse(fsshttpbResponse.DataElementPackage);
                            // Capture the MS-ONESTORE related requirements
                            new MsonestoreCapture().Validate(package, site);
                        }
                    }

                    if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                    {
                        new MsfsshttpbAdapterCapture().VerifyTransport(site);

                        // Capture the response related requirements
                        new MsfsshttpbAdapterCapture().VerifyFsshttpbResponse(fsshttpbResponse, site);
                    }
                }
            }

            // Validating the fragment of the sub response
            // Record the validation errors and warnings.
            ValidationResult result = SchemaValidation.ValidateXml(subResponseDocument.OuterXml);

            if (!SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (result != ValidationResult.Success)
                {
                    // Add error log
                    site.Assert.Fail("Schema validation fails, the reason is " + SchemaValidation.GenerateValidationResult());
                }

                // No need to run the capture code, just return.
                return;
            }

            if (result == ValidationResult.Success)
            {
                // Capture the requirement related to the sub response token.
                MsfsshttpAdapterCapture.ValidateSubResponseToken(site);

                // Call corresponding sub response capture code.
                this.InvokeCaptureCode(subResponse, site);
            }
            else
            {
                // Add error log
                site.Assert.Fail("Schema validation fails, the reason is " + SchemaValidation.GenerateValidationResult());
            }
        }

        /// <summary>
        /// This method is used to serialize the sub response to a specified sub response instance.
        /// </summary>
        /// <param name="subResponseDocument">Specify the sub response XML document.</param>
        /// <param name="site">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        /// <returns>Return the object represents the specified kind of sub response.</returns>
        private object SerializeSubResponse(XmlDocument subResponseDocument, ITestSite site)
        {
            Assembly assembly = Assembly.LoadFrom("Common.dll");
            if (assembly == null)
            {
                site.Assert.Fail("Cannot load the common object assembly.");
            }

            Type subResponseType = assembly.GetType("Microsoft.Protocols.TestSuites.Common." + this.SubResponseTypeName);
            if (subResponseType == null)
            {
                site.Assert.Fail(string.Format("Cannot load the type {0} from the assembly CommonProject.dll", this.SubResponseTypeName));
            }

            XmlAttributes xmlAttrs = new XmlAttributes();
            xmlAttrs.XmlType = new XmlTypeAttribute(this.SubResponseElementName);

            XmlAttributeOverrides xmlOverrides = new XmlAttributeOverrides();
            xmlOverrides.Add(subResponseType, xmlAttrs);

            XmlReflectionImporter importer = new XmlReflectionImporter(xmlOverrides, "http://schemas.microsoft.com/sharepoint/soap/");
            XmlSerializer serializer = new XmlSerializer(importer.ImportTypeMapping(subResponseType));

            using (MemoryStream ms = new MemoryStream())
            {
                byte[] subResponseBytes = System.Text.Encoding.UTF8.GetBytes(subResponseDocument.OuterXml);
                ms.Write(subResponseBytes, 0, subResponseBytes.Length);
                ms.Position = 0;
                return serializer.Deserialize(ms);
            }
        }

        /// <summary>
        /// This method is used to invoke capture logic according the sub response type.
        /// </summary>
        /// <param name="subResponse">Specify the sub response object instance.</param>
        /// <param name="site">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        private void InvokeCaptureCode(object subResponse, ITestSite site)
        {
            // Find the capture function 
            MethodInfo captureMethod = typeof(MsfsshttpAdapterCapture).GetMethod("Validate" + this.SubResponseElementName);
            if (captureMethod == null)
            {
                throw new InvalidOperationException(string.Format("Cannot find the function Validate{0} in the type MsFsshttpAdapterCapture.", this.SubResponseElementName));
            }

            captureMethod.Invoke(null, new object[] { subResponse, site });
        }

        /// <summary>
        /// This method is used to get the corresponding sub response according to the current record sub request token.
        /// </summary>
        /// <param name="rawResponse">Specify the raw XML string response returned by the protocol server.</param>
        /// <returns>Return the XmlDocument instance represents the sub response replaced with the SubResponseElementName.</returns>
        private XmlDocument ExtractSubResponseNode(string rawResponse)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(rawResponse);
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("ns", "http://schemas.microsoft.com/sharepoint/soap/");
            XmlNodeList subResponseNodes = doc.SelectNodes("//ns:SubResponse", nsmgr);

            foreach (XmlNode subResponse in subResponseNodes)
            {
                XmlAttribute tokenAttribute = subResponse.Attributes["SubRequestToken"];
                if (tokenAttribute == null)
                {
                    throw new System.InvalidOperationException("The SubRequestToken attribute must exist in the sub response element.");
                }

                if (string.Compare(this.SubToken, tokenAttribute.Value, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    XmlDocument newDocument = new XmlDocument();
                    XmlDeclaration declare = newDocument.CreateXmlDeclaration("1.0", "UTF-8", "yes");
                    newDocument.AppendChild(declare);

                    // Create the element
                    XmlNode newEle = doc.CreateNode(XmlNodeType.Element, this.SubResponseElementName, "http://schemas.microsoft.com/sharepoint/soap/");

                    // Clone old attribute for new node.
                    foreach (XmlAttribute subResAttr in subResponse.Attributes)
                    {
                        newEle.Attributes.Append(subResAttr.CloneNode(true) as XmlAttribute);
                    }

                    // Clone old subNode for new node.
                    XmlNodeList subResDataList = subResponse.ChildNodes;
                    if (subResDataList.Count != 0)
                    {
                        foreach (XmlNode subResDatas in subResDataList)
                        {
                            newEle.AppendChild(subResDatas.CloneNode(true));
                        }
                    }

                    XmlNode newImportNode = newDocument.ImportNode(newEle, true);
                    newDocument.AppendChild(newImportNode);
                    return newDocument;
                }
            }

            throw new System.InvalidOperationException("Cannot get the sub response using the sub request token" + this.SubToken);
        }
    }
}