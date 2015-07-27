//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// A class is used to perform serializer operations for MS-WOPI protocol.
    /// </summary>
    public class WOPIResponseHelper : HelperBase
    {
        /// <summary>
        /// Prevents a default instance of the WOPIResponseHelper class from being created
        /// </summary>
        private WOPIResponseHelper()
        {
        }

        /// <summary>
        /// A method is used to read the HTTP response body to the bytes array.
        /// </summary>
        /// <param name="wopiHttpResponse">A parameter represents the HTTP response.</param>
        /// <returns>A return value represents the raw body content. If the body length is larger than (int.MaxValue) bytes, the body contents will be chunked by 1024 bytes. The max length of this method is (1024 * int.MaxValue) bytes.</returns>
        public static List<byte[]> ReadRawHTTPResponseToBytes(WOPIHttpResponse wopiHttpResponse)
        {
            if (null == wopiHttpResponse)
            {
                throw new ArgumentNullException("wopiHttpResponse");
            }

            using (Stream bodyStream = wopiHttpResponse.GetResponseStream())
            {
                long contentLengthValue = wopiHttpResponse.ContentLength;
                return ReadBytesFromHttpBodyStream(bodyStream, contentLengthValue);
            }
        }

        /// <summary>
        /// A method is used to read the HTTP response body and decode to string.
        /// </summary>
        /// <param name="wopiHttpResponse">A parameter represents the HTTP response.</param>
        /// <returns>A return value represents the string which is decode from the HTTP response body. The decode format is UTF-8 by default.</returns>
        public static string ReadHTTPResponseBodyToString(WOPIHttpResponse wopiHttpResponse)
        {
            if (null == wopiHttpResponse)
            {
                throw new ArgumentNullException("wopiHttpResponse");
            }

            string bodyString = string.Empty;
            long bodyLength = wopiHttpResponse.ContentLength;
            if (bodyLength != 0)
            {
                Stream bodStream = null;
                try
                {
                    bodStream = wopiHttpResponse.GetResponseStream();
                    using (StreamReader strReader = new StreamReader(bodStream))
                    {
                        bodyString = strReader.ReadToEnd();
                    }
                }
                finally
                {
                    if (bodStream != null)
                    {
                        bodStream.Dispose();
                    }
                }
            }

            return bodyString;
        }

        /// <summary>
        /// A method is used to get raw body contents whose length is in 1 to int.MaxValue bytes scope.
        /// </summary>
        /// <param name="wopiHttpResponse">A parameter represents the HTTP response.</param>
        /// <returns>A return value represents the raw body content.</returns>
        public static byte[] GetContentFromResponse(WOPIHttpResponse wopiHttpResponse)
        {
            List<byte[]> rawBytesOfBody = ReadRawHTTPResponseToBytes(wopiHttpResponse);
            byte[] returnContent = rawBytesOfBody.SelectMany(bytes => bytes).ToArray();

            DiscoveryProcessHelper.AppendLogs(
                      typeof(WOPIResponseHelper),
                      string.Format(
                                "Read normal size(1 to int.MaxValue bytes) data from response. actual size[{0}] bytes",
                                 returnContent.Length));

            return returnContent;
        }

        /// <summary>
        /// A method used to get byte string values from the response. 50 bytes per line.
        /// </summary>
        /// <param name="rawData">A parameter represents a WOPI response instance which contains the response body.</param>
        /// <returns>A return value represents byte string values.</returns>
        public static string GetBytesStringValue(byte[] rawData)
        {
            if (null == rawData || 0 == rawData.Length)
            {
                throw new ArgumentNullException("rawData");
            }

            string bitsString = BitConverter.ToString(rawData);
            char[] splitSymbol = new char[] { '-' };
            string[] bitStringValue = bitsString.Split(splitSymbol, StringSplitOptions.RemoveEmptyEntries);

            // Output 50 bits string value per line.
            int wrapNumberOfBitString = 50;
            if (bitStringValue.Length <= wrapNumberOfBitString)
            {
                return bitsString;
            }
            else
            {
                StringBuilder strBuilder = new StringBuilder();
                int beginIndexOfOneLine = 0;
                bool stopSplitLine = false;
                int oneLineLength = wrapNumberOfBitString * 3;

                // Wrap the bits string value.
                do
                {
                    // If it meets the last line, get the actual bits value.
                    if (beginIndexOfOneLine + oneLineLength > bitsString.Length)
                    {
                        oneLineLength = bitsString.Length - beginIndexOfOneLine;
                        stopSplitLine = true;
                    }

                    string oneLineContent = bitsString.Substring(beginIndexOfOneLine, oneLineLength);
                    beginIndexOfOneLine = beginIndexOfOneLine + oneLineLength;
                    strBuilder.AppendLine(oneLineContent);
                }
                while (!stopSplitLine);

                return strBuilder.ToString();
            }
        }

        /// <summary>
        /// A method used to read bytes data from the stream of the HTTP response.
        /// </summary>
        /// <param name="bodyBinaries">A parameter represents the stream which contain the body binaries data.</param>
        /// <param name="contentLengthValue">A parameter represents the length of the body binaries.</param>
        /// <returns>A return value represents the raw body content. If the body length is larger than (int.MaxValue) bytes, the body contents will be chunked by 1024 bytes. The max length of this method is (1024 * int.MaxValue) bytes.</returns>
        private static List<byte[]> ReadBytesFromHttpBodyStream(Stream bodyBinaries, long contentLengthValue)
        {
            if (null == bodyBinaries)
            {
                throw new ArgumentNullException("bodyBinaries");
            }

            long maxKBSize = (long)int.MaxValue;
            long totalSize = contentLengthValue;
            if (contentLengthValue > maxKBSize * 1024)
            {
                throw new InvalidOperationException(string.Format("The test suite only support [{0}]KB size content in a HTTP response.", maxKBSize));
            }

            List<byte[]> bytesOfResponseBody = new List<byte[]>();
            bool isuseChunk = false;
            using (BinaryReader binReader = new BinaryReader(bodyBinaries))
            {
                if (contentLengthValue < int.MaxValue && contentLengthValue > 0)
                {
                    byte[] bytesBlob = binReader.ReadBytes((int)contentLengthValue);
                    bytesOfResponseBody.Add(bytesBlob);
                }
                else
                {
                    isuseChunk = true;
                    totalSize = 0;

                    // set chunk size to 1KB per reading.
                    int chunkSize = 1024;
                    byte[] bytesBlockTemp = null;

                    do
                    {
                        bytesBlockTemp = binReader.ReadBytes(chunkSize);
                        if (bytesBlockTemp.Length > 0)
                        {
                            bytesOfResponseBody.Add(bytesBlockTemp);
                            totalSize += bytesBlockTemp.Length;
                        }
                    }
                    while (bytesBlockTemp.Length > 0);
                }
            }

            string chunkProcessLogs = string.Format(
                             "Read [{0}] KB size data from response body.{1}",
                             (totalSize / 1024.0).ToString("f2"),
                             isuseChunk ? "Use the chunk reading mode." : string.Empty);
            DiscoveryProcessHelper.AppendLogs(typeof(WOPIResponseHelper), chunkProcessLogs);

            return bytesOfResponseBody;
        }
    }
}