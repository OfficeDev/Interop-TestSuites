namespace Microsoft.Protocols.TestSuites.MS_WWSP
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Xml;
    using System.Xml.Schema;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This class contains common help methods.
    /// </summary>
    public static class AdapterHelper
    {
        #region Variables

        /// <summary>
        /// Used to read configuration property from PTF configuration and capture requirements.
        /// </summary>
        private static ITestSite site;

        #endregion Variables

        #region Adapter help methods

        /// <summary>
        /// Initialize object of "Site".
        /// </summary>
        /// <param name="currentSite">An object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        public static void Initialize(ITestSite currentSite)
        {
            site = currentSite;
        }

        /// <summary>
        /// Get attribute value from response data of an operation.
        /// </summary>
        /// <param name="attributeName">The specified attribute name.</param>
        /// <param name="xmlNode">The response data of an operation.</param>
        /// <returns>The attribute value.</returns>
        public static string GetAttributeValueFromXml(string attributeName, XmlNode xmlNode)
        {
            string attributeValue = null;
            if (xmlNode.Name.Equals("ToDoData"))
            {
                try
                {
                    // The structure of this statement is decided by the template data structure.
                    attributeValue = xmlNode.FirstChild.FirstChild.FirstChild.Attributes.GetNamedItem(attributeName).Value;
                }
                catch
                {
                    // If can't find the expected XML node, will be caught here and set the value to null.
                    attributeValue = null;
                }
            }
            else
            {
                try
                {
                    attributeValue = xmlNode.Attributes.GetNamedItem(attributeName).Value;
                }
                catch
                {
                    // If can't find the expected XML node, will be caught here and set the value to null.
                    attributeValue = null;
                }
            }

            return attributeValue;
        }

        /// <summary>
        /// Get the specified XML node from the specified XML fragment.
        /// </summary>
        /// <param name="nodeName">The name of specified XML node.</param>
        /// <param name="xmlNode">The specified XML fragment.</param>
        /// <returns>The result of the research.</returns>
        public static XmlNode GetNodeFromXML(string nodeName, XmlNode xmlNode)
        {
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlNode.OwnerDocument.NameTable);
            nsmgr.AddNamespace("wf", "http://schemas.microsoft.com/sharepoint/soap/workflow/");
            XmlNode elementNode = xmlNode.SelectSingleNode("//wf:" + nodeName, nsmgr);

            site.Assert.IsNotNull(elementNode, "The element Node should not be null.");

            return elementNode;
        }

        /// <summary>
        /// The method is used to verify whether specified elementName is existed in the specified lastRawXml.
        /// </summary>
        /// <param name="xmlElement">The XML element.</param>
        /// <param name="elementName">The element name which need to check whether it is existed.</param>
        /// <returns>If the XML response has contain element, true means include, otherwise false.</returns>
        public static bool HasElement(XmlElement xmlElement, string elementName)
        {
            // Verify whether elementName is existed.
            // If server response XML contains elementName, true will be returned. otherwise false will be returned.
            if (xmlElement.GetElementsByTagName(elementName).Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Verify the schema definition for the workflow association data, if there is error in validation, method will throw an XmlSchemaValidationException.
        /// </summary>
        /// <param name="workFlowAssociationData">A parameter represents the data of the association</param>
        public static void VerifyWorkflowAssociationSchema(XmlNode workFlowAssociationData)
        {
            if (null == workFlowAssociationData || string.IsNullOrEmpty(workFlowAssociationData.OuterXml))
            {
                throw new ArgumentException("The [workFlowAssociationData] parameter should be have instance and contain valid OuterXml");
            }

            // Validate the WorkFlowAssociation schema for TemplateData.WorkflowTemplates.WorkflowTemplate.AssocationData.
            bool enableworkFlowAssociationValidation = Common.GetConfigurationPropertyValue<bool>("ValidateWorkFlowAssociation", site);
            string contents = string.Empty;
            if (enableworkFlowAssociationValidation)
            {
                // WorkFlowAssociationData element might contain an inner text which is the association content or contain a child element which includes the association content.
                if (!workFlowAssociationData.HasChildNodes)
                {
                    contents = workFlowAssociationData.InnerText;
                }
                else
                {
                    contents = workFlowAssociationData.FirstChild.InnerText;
                }

                if (!string.IsNullOrEmpty(contents))
                {
                    List<string> schemaDefinitions = LoadWorkflowAssociationSchemas();

                    // Get schema definitions from xsd file.
                    ValidationResult validationResult = XmlValidator.ValidateXml(schemaDefinitions, contents);

                    // If have validation error, throw new validation exception
                    if (validationResult != ValidationResult.Success)
                    {
                        throw new XmlSchemaValidationException(
                                 string.Format(
                                     "There are schema validation errors or warnings when validating the Workflow association data in response of GetTemplatesForItem operation, the result is {0}",
                                     XmlValidator.GenerateValidationResult()));
                    }
                }
            }
        }

        /// <summary>
        /// A method used to check if the specified value is zero or more combination of the bitmasks.
        /// </summary>
        /// <param name="value">A parameter represents the value which is used to check.</param>
        /// <param name="bitMasks">A parameter represents the array which is used limit the value specified in value parameter</param>
        /// <returns>Returns true indicating the value is valid.</returns>
        public static bool IsValueValid(long value, long[] bitMasks)
        {
            // Retrieve all the combinations(more bitmasks).
            List<long> combinations = GetCombinationsFromBitMasks(bitMasks);

            // Add zero to the possible values.
            combinations.Add(0);

            return combinations.Contains(value);
        }

        /// <summary>
        /// Get the WorkFlow Template Item By specified Name
        /// </summary>
        /// <param name="templateName">A parameter represents the template name which will be used to find out the template item</param>
        /// <param name="templateData">A parameter represents response of GetTemplatesForItem operation.</param>
        /// <returns>A return represents the template item data.</returns>
        public static TemplateDataWorkflowTemplate GetWorkFlowTemplateItemByName(string templateName, TemplateData templateData)
        {
            if (string.IsNullOrEmpty(templateName) || null == templateData)
            {  
               string errMsg = string.Format(
                   "All Parameters should not be null or empty: templateName[{0}] getTemplatesForItemResult[{1}]",
                   string.IsNullOrEmpty(templateName) ? "NullOrEmpty" : "Valid",
                   null == templateData ? "Null" : "Valid");
               throw new ArgumentException(errMsg);
            }

            if (null == templateData.WorkflowTemplates)
            {
               site.Assert.Fail("Could not get the valid TemplateData from the response of GetTemplatesForItem operation.");
            }

            TemplateDataWorkflowTemplate[] templates = templateData.WorkflowTemplates;

            var expectedTemplateItems = from templateItem in templates
                                        where templateItem.Name.Equals(templateName, StringComparison.OrdinalIgnoreCase)
                                        select templateItem;

            TemplateDataWorkflowTemplate matchTemplateItem = null;
            int itemsCounter = expectedTemplateItems.Count();
            if (1 < itemsCounter)
            {
                site.Assert.Fail("The response of GetTemplatesForItem operation should contain only one matched TemplateData item.");
            }
            else if (0 == itemsCounter)
            {
                return matchTemplateItem;
            }
            else
            {
                matchTemplateItem = expectedTemplateItems.ElementAt(0);
            }

            return matchTemplateItem;
        }

        /// <summary>
        /// Get the association data from specified templateItem in response of GetTemplatesForItem operation.
        /// </summary>
        /// <param name="templateName">A parameter represents the template name which will be used to find out the template item</param>
        /// <param name="templateData">A parameter represents response of GetTemplatesForItem operation which contains the association data.</param>
        /// <returns>A return represents the association data.</returns>
        public static XmlNode GetAssociationDataFromTemplateItem(string templateName, TemplateData templateData)
        {
            TemplateDataWorkflowTemplate currentWorkflowTemplateItem = GetWorkFlowTemplateItemByName(templateName, templateData);
            if (null == currentWorkflowTemplateItem)
            {
                site.Assert.Fail(
                            "The response of getTemplatesForItem operation should contain template item with expected name[{0}]",
                            templateName);
            }

            return currentWorkflowTemplateItem.AssociationData;
        }

        #endregion Adapter help methods

        #region Extend methods

        /// <summary>
        /// It is extend method and used to compare fields' value between two instance of ClaimReleaseTaskResponseClaimReleaseTaskResult type. 
        /// </summary>
        /// <param name="currentclaimResultInstance">A parameter represents the current instance of ClaimReleaseTaskResponseClaimReleaseTaskResult type.</param>
        /// <param name="targetclaimResultInstance">A parameter represents the target instance of ClaimReleaseTaskResponseClaimReleaseTaskResult type which will be compared.</param>
        /// <returns>Return true indicating the current claimResultInstance is equal to target claimResultInstance.</returns>
        public static bool AreEquals(this ClaimReleaseTaskResponseClaimReleaseTaskResult currentclaimResultInstance, ClaimReleaseTaskResponseClaimReleaseTaskResult targetclaimResultInstance)
        {
            if (null == targetclaimResultInstance)
            {
                return false;
            }

            ClaimReleaseTaskResponseClaimReleaseTaskResultTaskData currentTaskData = currentclaimResultInstance.TaskData;
            ClaimReleaseTaskResponseClaimReleaseTaskResultTaskData targetTaskData = targetclaimResultInstance.TaskData;
            bool compareResult = string.Equals(currentTaskData.AssignedTo, targetTaskData.AssignedTo, StringComparison.OrdinalIgnoreCase);
            compareResult = compareResult && int.Equals(currentTaskData.ItemId, targetTaskData.ItemId);
            compareResult = compareResult && Guid.Equals(currentTaskData.ListId, targetTaskData.ListId);
            compareResult = compareResult && string.Equals(currentTaskData.TaskGroup, targetTaskData.TaskGroup);
            return compareResult;
        }
        
        #endregion Extend methods

        #region private methods

        /// <summary>
        /// Load the workflow Association Schema definitions
        /// </summary>
        /// <returns>A return represents the schema definitions of workflow association</returns>
        private static List<string> LoadWorkflowAssociationSchemas()
        {
            string workflowAssociationSchemaFile = Common.GetConfigurationPropertyValue("WorkFlowAssociationXsdFile", site);
            if (string.IsNullOrEmpty(workflowAssociationSchemaFile))
            {
                throw new Exception("The workflowAssociationSchemaFile property value should not be empty when enable the Association data schema validation.");
            }

            // Process the workflowAssociation SchemaFile for different SUT in Microsoft Products
            if (workflowAssociationSchemaFile.IndexOf(@"[SUTVersionShortName]", StringComparison.OrdinalIgnoreCase) > 0)
            {
                workflowAssociationSchemaFile = workflowAssociationSchemaFile.ToLower();
                string expectedSutPlaceHolderValue = string.Empty;
                string currentVersion = Common.GetConfigurationPropertyValue("SUTVersion", site);

                if (currentVersion.Equals("SharePointServer2007", StringComparison.OrdinalIgnoreCase))
                {
                    expectedSutPlaceHolderValue = "2007";
                }
                else if (currentVersion.Equals("SharePointServer2010", StringComparison.OrdinalIgnoreCase))
                {
                    expectedSutPlaceHolderValue = "2010";
                }
                else if (currentVersion.Equals("SharePointServer2013", StringComparison.OrdinalIgnoreCase))
                {
                    expectedSutPlaceHolderValue = "2013";
                }
                else if (currentVersion.Equals("SharePointServer2016", StringComparison.OrdinalIgnoreCase))
                {
                    expectedSutPlaceHolderValue = "2016";
                }
                else if (currentVersion.Equals("SharePointServer2019", StringComparison.OrdinalIgnoreCase))
                {
                    expectedSutPlaceHolderValue = "2019";
                }
                else if (currentVersion.Equals("SharePointServerSubscriptionEdition", StringComparison.OrdinalIgnoreCase))
                {
                    expectedSutPlaceHolderValue = "SubscriptionEdition";
                }
                else
                {
                    throw new Exception("Could Not Generate correct workflowAssociation Schema File name.");
                }

                workflowAssociationSchemaFile = workflowAssociationSchemaFile.Replace("[SUTVersionShortName]".ToLower(), expectedSutPlaceHolderValue);
            }

            #region Process multiple schema definitions in one file.
            XmlDocument doc = new XmlDocument();
            doc.Load(workflowAssociationSchemaFile);

            XmlElement rootElement = doc.DocumentElement;
            List<string> schemaDefinitions = new List<string>();

            // if it is single "Schema definition" in this file.
            if (rootElement.LocalName.Equals("schema", StringComparison.OrdinalIgnoreCase))
            {
                schemaDefinitions.Add(rootElement.OuterXml);
                return schemaDefinitions;
            }

            // multiple  "Schema definitions" in this file, and test suite will use "SchemaXsds" xml element to contain multiple definitions.
            if (!rootElement.LocalName.Equals("SchemaXsds", StringComparison.OrdinalIgnoreCase))
            {
                throw new Exception("The workflow association schema definition file's root element should be [SchemaXsds] or [schema].");
            }

            if (!rootElement.HasChildNodes)
            {
                throw new Exception("The workflow association schema definition file should contain at least one schema definition under [SchemaXsds] element.");
            }

            var validSchemaDefinitionNode = from XmlNode schemaItem in rootElement.ChildNodes
                                    where schemaItem.LocalName.Equals("schema", StringComparison.OrdinalIgnoreCase)
                                    select schemaItem;

            foreach (XmlNode schemadefinition in validSchemaDefinitionNode)
            {
                schemaDefinitions.Add(schemadefinition.OuterXml);
            }

            return schemaDefinitions;
            #endregion
        }

        /// <summary>
        /// Get all the combinations of the bitmasks.
        /// </summary>
        /// <param name="bitMasks">The bitmask array.</param>
        /// <returns>All the combinations of the bitmasks.</returns>
        private static List<long> GetCombinationsFromBitMasks(long[] bitMasks)
        {
            List<long> combinations = new List<long>();
            int countOfBitMasks = bitMasks.Length;

            for (int index = 1; index <= countOfBitMasks; ++index)
            {
                Combination(bitMasks, index, ref combinations);
            }

            return combinations;
        }

        /// <summary>
        /// Get the combinations from the specified count of bitmasks.
        /// </summary>
        /// <param name="bitMasks">The bitmask array.</param>
        /// <param name="countCom">The count of bitmasks the combinations contains.</param>
        /// <param name="combinations">The combination values.</param>
        /// <returns>"true": success; "false": failed.</returns>
        private static bool Combination(long[] bitMasks, int countCom, ref List<long> combinations)
        {
            int length = bitMasks.Length;
            if (length < countCom)
            {
                return false;
            }

            long[] array = new long[length];
            long indexFirst = 0;

            // Initialize array.
            for (indexFirst = 0; indexFirst < length; indexFirst++)
            {
                array[indexFirst] = 0;
            }

            // Calculate possible bitmask values and add them into "combinations".
            long indexSecond = 0;
            while (indexSecond >= 0)
            {
                if (array[indexSecond] < (length - countCom + indexSecond + 1))
                {
                    indexFirst = indexSecond;
                    array[indexSecond]++;
                }
                else
                {
                    indexSecond--;
                    continue;
                }

                for (; indexFirst < countCom - 1; indexFirst++)
                {
                    array[indexFirst + 1] = array[indexFirst] + 1;
                }

                if (indexFirst == countCom - 1)
                {
                    long result = 0L;
                    for (int idxBit = 0; idxBit < countCom; ++idxBit)
                    {
                        result += bitMasks[array[idxBit] - 1];
                    }

                    combinations.Add(result);
                }

                indexSecond = indexFirst;
            }

            return true;
        }
        
        #endregion
    }
}