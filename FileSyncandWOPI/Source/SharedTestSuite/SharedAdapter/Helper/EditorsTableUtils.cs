namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;
    using System.IO.Compression;
    using System.Xml;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This method is used to extract EditorsTable from the MS-FSSHTTPB binary response.
    /// </summary>
    public static class EditorsTableUtils
    {
        /// <summary>
        /// Static field for the editors table header byte content.
        /// </summary>
        public static readonly byte[] EditorsTableHeader = new byte[] { 0x1a, 0x5a, 0x3a, 0x30, 0, 0, 0, 0 };

        /// <summary>
        /// This method is used to test whether it is an editors table header.
        /// </summary>
        /// <param name="content">Specify the header content.</param>
        /// <returns>Return true if the content is editors table header, otherwise return false.</returns>
        public static bool IsEditorsTableHeader(byte[] content)
        {
            byte[] editorsTableHeaderTmp = new byte[8];

            if (content.Length < 8)
            {
                return false;
            }

            Array.Copy(content, 0, editorsTableHeaderTmp, 0, 8);
            return AdapterHelper.ByteArrayEquals(EditorsTableHeader, editorsTableHeaderTmp);
        }

        /// <summary>
        /// Get EditorsTable from server response.
        /// </summary>
        /// <param name="subResponse">The sub response from server.</param>
        /// <param name="site">Transfer ITestSite into this operation, for this operation to use ITestSite's function.</param>
        /// <returns>The instance of EditorsTable.</returns>
        public static EditorsTable GetEditorsTableFromResponse(FsshttpbResponse subResponse, ITestSite site)
        {
            if (subResponse == null || subResponse.DataElementPackage == null || subResponse.DataElementPackage.DataElements == null)
            {
                site.Assert.Fail("The parameter CellResponse is not valid, check whether the CellResponse::DataElementPackage or CellResponse::DataElementPackage::DataElements is null.");
            }

            foreach (DataElement de in subResponse.DataElementPackage.DataElements)
            {
                if (de.Data.GetType() == typeof(ObjectGroupDataElementData))
                {
                    ObjectGroupDataElementData ogde = de.Data as ObjectGroupDataElementData;

                    if (ogde.ObjectGroupData == null || ogde.ObjectGroupData.ObjectGroupObjectDataList.Count == 0)
                    {
                        continue;
                    }

                    for (int i = 0; i < ogde.ObjectGroupData.ObjectGroupObjectDataList.Count; i++)
                    {
                        if (IsEditorsTableHeader(ogde.ObjectGroupData.ObjectGroupObjectDataList[i].Data.Content.ToArray()))
                        {
                            string editorsTableXml = null;

                            // If the current object group object data is the header byte array 0x1a, 0x5a, 0x3a, 0x30, 0, 0, 0, 0, then the immediate following object group object data will contain the Editor table xml. 
                            byte[] buffer = ogde.ObjectGroupData.ObjectGroupObjectDataList[i + 1].Data.Content.ToArray();
                            System.IO.MemoryStream ms = null;
                            try
                            {
                                ms = new System.IO.MemoryStream();
                                ms.Write(buffer, 0, buffer.Length);
                                ms.Position = 0;
                                using (DeflateStream stream = new DeflateStream(ms, CompressionMode.Decompress))
                                {
                                    stream.Flush();
                                    byte[] decompressBuffer = new byte[buffer.Length * 3];
                                    int decompressdSize = stream.Read(decompressBuffer, 0, buffer.Length * 3);
                                    Array.Resize(ref decompressBuffer, decompressdSize);
                                    stream.Close();
                                    editorsTableXml = System.Text.Encoding.UTF8.GetString(decompressBuffer);
                                }

                                ms.Close();
                            }
                            finally
                            {
                                if (ms != null)
                                {
                                    ms.Dispose();
                                }
                            }

                            return GetEditorsTable(editorsTableXml);
                        }
                    }
                }
            }

            throw new InvalidOperationException("Cannot find any data group object data contain editor tables information.");
        }

        /// <summary>
        /// Get EditorsTable from response xml.
        /// </summary>
        /// <param name="responseXml">The response xml about EditorsTable.</param>
        /// <returns>The instance of EditorsTable.</returns>
        public static EditorsTable GetEditorsTable(string responseXml)
        {
            responseXml = System.Text.RegularExpressions.Regex.Replace(responseXml, "^[^<]", string.Empty);
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(responseXml);
            XmlNodeList nodeList = doc.GetElementsByTagName("Editor");
            List<Editor> list = new List<Editor>();
            if (nodeList.Count > 0)
            {
                foreach (XmlNode node in nodeList)
                {
                    list.Add(GetEditor(node));
                }
            }

            EditorsTable table = new EditorsTable();
            table.Editors = list.ToArray();

            return table;
        }

        /// <summary>
        /// Get Editor instance from XmlNode.
        /// </summary>
        /// <param name="node">The XmlNode which contents the Editor data.</param>
        /// <returns>Then instance of Editor.</returns>
        private static Editor GetEditor(XmlNode node)
        {
            if (node.ChildNodes.Count == 0)
            {
                return null;
            }

            Editor editors = new Editor();
            foreach (XmlNode item in node.ChildNodes)
            {
                object propValue;
                if (item.Name == "HasEditorPermission")
                {
                    propValue = Convert.ToBoolean(item.InnerText);
                }
                else if (item.Name == "Timeout")
                {
                    propValue = Convert.ToInt64(item.InnerText);
                }
                else if (item.Name == "Metadata")
                {
                    Dictionary<string, string> metaData = new Dictionary<string, string>();

                    foreach (XmlNode metaNode in item.ChildNodes)
                    {
                        metaData.Add(metaNode.Name, new System.Text.UnicodeEncoding().GetString(Convert.FromBase64String(metaNode.InnerText)));
                    }

                    propValue = metaData;
                }
                else
                {
                    propValue = item.InnerText;
                }

                FsshttpConverter.SetSpecifiedProtyValueByName(editors, item.Name, propValue);
            }

            return editors;
        }
    }
}