//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// A chunked response for Mailbox Server Endpoint.
    /// </summary>
    public class ChunkedResponse
    {
        /// <summary>
        /// Gets the MetaTags property
        /// </summary>
        public List<string> MetaTags { get; private set; }

        /// <summary>
        /// Gets the AdditionHeaders property
        /// </summary>
        public Dictionary<string, string> AdditionHeaders { get; private set; }

        /// <summary>
        /// Gets the ResponseBodyRawData property
        /// </summary>
        public byte[] ResponseBodyRawData { get; private set; }

        /// <summary>
        /// The method to parse the chunked response from server.
        /// </summary>
        /// <param name="rawData">The response data from the server.</param>
        /// <returns>The structure for chunked response.</returns>
        public static ChunkedResponse ParseChunkedResponse(byte[] rawData)
        {
            ChunkedResponse response = new ChunkedResponse();
            response.MetaTags = new List<string>();
            response.AdditionHeaders = new Dictionary<string, string>();
            List<byte> responseBodyData = new List<byte>();
            bool isDone = false;
            bool isStartReadBody = false;
            do
            {
                byte[] lineData = ReadLine(ref rawData, isStartReadBody);
                string line = Encoding.ASCII.GetString(lineData);
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
                            string[] headerkeyAndValue = line.Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                            response.AdditionHeaders.Add(headerkeyAndValue[0], headerkeyAndValue[1]);
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
        /// The method to read one line from the response data.
        /// </summary>
        /// <param name="currentData">The current data will be processed.</param>
        /// <param name="isKeepReturn">The flag to indicate if to keep the return key in the line.</param>
        /// <returns>The line will be returned.</returns>
        private static byte[] ReadLine(ref byte[] currentData, bool isKeepReturn)
        {
            int length = 0;
            byte[] lineBuffer;

            for (int i = 0; i < currentData.Length - 1; i++)
            {
                // To check the Return key in the currentData
                if (currentData[i] == 0x0D && currentData[i + 1] == 0x0A)
                {
                    if (isKeepReturn)
                    {
                        lineBuffer = new byte[length + 2];
                        Array.Copy(currentData, lineBuffer, length + 2);
                    }
                    else
                    {
                        lineBuffer = new byte[length];
                        Array.Copy(currentData, lineBuffer, length);
                    }

                    if (i + 2 < currentData.Length)
                    {
                        byte[] laveData = new byte[currentData.Length - length - 2];
                        Array.Copy(currentData, i + 2, laveData, 0, currentData.Length - length - 2);
                        currentData = laveData;
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
