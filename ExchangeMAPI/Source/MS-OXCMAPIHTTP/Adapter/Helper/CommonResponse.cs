//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// This class used for chunk response
    /// </summary>
    public class CommonResponse
    {
        /// <summary>
        /// Gets and set meta tags in the response.
        /// </summary>
        public List<string> MetaTags { get; private set; }

        /// <summary>
        /// Gets additional headers.
        /// </summary>
        public Dictionary<string, string> AdditionalHeaders { get; private set; }

        /// <summary>
        /// Gets the response body raw data.
        /// </summary>
        public byte[] ResponseBodyRawData { get; private set; }

        /// <summary>
        /// This method is used to parse the chunk response.
        /// </summary>
        /// <param name="rawData">The raw data be parsed</param>
        /// <returns>The parsed chunk response</returns>
        public static CommonResponse ParseCommonResponse(byte[] rawData)
        {        
            CommonResponse response = new CommonResponse();
            response.MetaTags = new List<string>();
            response.AdditionalHeaders = new Dictionary<string, string>();
            List<byte> responseBodyData = new List<byte>();
            bool isDone = false;
            bool isStartReadBody = false;
            do
            {
                byte[] lineData = ReadLine(ref rawData);
                string line = Encoding.ASCII.GetString(lineData).Replace("\r\n", string.Empty);
                if (isDone == false)
                {
                    if (string.Compare(line, "PROCESSING", true) == 0 || string.Compare(line, "PENDING", true) == 0)
                    {
                        response.MetaTags.Add(line);
                    }
                    else if (string.Compare(line, "DONE", true) == 0)
                    {
                        response.MetaTags.Add(line);
                        isDone = true;
                    }
                }
                else
                {
                    if (isStartReadBody == false)
                    {
                       if (line == string.Empty)
                       {
                           isStartReadBody = true;
                       }
                       else
                       {
                           string headerName = line.Substring(0, line.IndexOf(":"));
                           string headerValue = line.Substring(line.IndexOf(":") + 1);
                           response.AdditionalHeaders.Add(headerName, headerValue);
                       }
                   }
                   else
                   {
                       responseBodyData.AddRange(lineData);
                   }
                }
            }
            while (rawData.Length > 0);
            response.ResponseBodyRawData = responseBodyData.ToArray();

            return response;
        }

        /// <summary>
        /// This method is used to read a line of the data
        /// </summary>
        /// <param name="currentData">The current data</param>
        /// <returns>The bytes of a line</returns>
        private static byte[] ReadLine(ref byte[] currentData)
        {
            int length = 0;
            byte[] lineBuffer;

            for (int i = 0; i < currentData.Length - 1; i++)
            {
                if (currentData[i] == 0x0D && currentData[i + 1] == 0x0A)
                {
                    lineBuffer = new byte[length + 2];
                    Array.Copy(currentData, lineBuffer, length + 2);

                    if (i + 2 < currentData.Length)
                    {
                        byte[] laveData = new byte[currentData.Length - length - 2];
                        Array.Copy(currentData, i + 2, laveData, 0, currentData.Length - length - 2);
                        currentData = laveData;
                    }
                    else
                    {
                        currentData = new byte[] { };
                    }

                    return lineBuffer;
                }
                else
                {
                    length += 1;
                }
            }

            lineBuffer = new byte[currentData.Length];
            Array.Copy(currentData, lineBuffer, currentData.Length);
            currentData = new byte[0] { };

            return lineBuffer;
        }
    }
}