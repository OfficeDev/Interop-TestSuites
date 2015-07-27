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
    using System.IO;
    using System.Net;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class that implements MapiHttp transport.
    /// </summary>
    public class MapiHttpAdapter
    {
        /// <summary>
        /// Cookie container for client
        /// </summary>
        private CookieCollection cookies = new CookieCollection();

        /// <summary>
        /// site used to write log to log files
        /// </summary>
        private ITestSite site;
        
        /// <summary>
        /// The name of the server to be connected
        /// </summary>
        private string mailStoreUrl = null;

        /// <summary>
        /// The domain of the user
        /// </summary>
        private string domain = null;

        /// <summary>
        /// The name of the user to connect to the server
        /// </summary>
        private string userName = null;

        /// <summary>
        /// The password of the user
        /// </summary>
        private string userPassword = null;

        /// <summary>
        /// Initializes a new instance of the MapiHttpAdapter class.
        /// </summary>
        /// <param name="site">site used to write log to log files</param>
        public MapiHttpAdapter(ITestSite site)
        {
            this.site = site;
        }
        
        /// <summary>
        /// The method to send MAPIHTTP request to the server.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
        /// <param name="mailStoreUrl">Mail store url.</param>
        /// <param name="userName">The user which connects the server.</param>
        /// <param name="domain">The domain of the user.</param>
        /// <param name="password">The password for the user.</param>
        /// <param name="requestBody">The MAPIHTTP request body.</param>
        /// <param name="requestType">The MAPIHTTP request type.</param>
        /// <param name="cookies">Cookie container for client.</param>
        /// <returns>Return the MAPIHTTP response from the server.</returns>
        public static HttpWebResponse SendMAPIHttpRequest(ITestSite site, string mailStoreUrl, string userName, string domain, string password, IRequestBody requestBody, string requestType, CookieCollection cookies)
        {
            HttpWebResponse response = null;

            System.Net.ServicePointManager.ServerCertificateValidationCallback =
            new System.Net.Security.RemoteCertificateValidationCallback(Common.ValidateServerCertificate);
            HttpWebRequest request = WebRequest.Create(mailStoreUrl) as HttpWebRequest;
            request.KeepAlive = true;
            request.CookieContainer = new CookieContainer();
            request.Method = "POST";
            request.ProtocolVersion = HttpVersion.Version11;
            request.Credentials = new System.Net.NetworkCredential(userName, password, domain);
            request.ContentType = "application/mapi-http";
            request.Accept = "application/mapi-http";
            request.Connection = string.Empty;

            byte[] buffer = null;
            if (requestBody != null)
            {
                buffer = requestBody.Serialize();
                request.ContentLength = buffer.Length;
            }
            else
            {
                request.ContentLength = 0;
            }

            request.Headers.Add("X-ClientInfo", "{A7A47AAD-233C-412B-9D10-DDE9108FEBD7}-5");
            request.Headers.Add("X-RequestId", "{16AC2587-EED8-48EB-8A7B-D48558B68BD7}:1");
            request.Headers.Add("X-ClientApplication", "Outlook/15.00.0856.000");
            request.Headers.Add("X-RequestType", requestType);
            if (cookies != null && cookies.Count > 0)
            {
                foreach (Cookie cookie in cookies)
                {
                    request.CookieContainer.Add(cookie);
                }
            }

            if (requestBody != null)
            {
                using (Stream stream = request.GetRequestStream())
                {
                    stream.Write(buffer, 0, buffer.Length);
                }
            }

            try
            {
                response = request.GetResponse() as HttpWebResponse;
            }
            catch (WebException ex)
            {
                site.Assert.Fail("A WebException happened when connecting the server, The exception is {0}.", ex.Message);
            }

            return response;
        }
        
        /// <summary>
        /// The method gets the binary data from the response data.
        /// </summary>
        /// <param name="response">The structure of the response data.</param>
        /// <returns>Returns the binary format of response data.</returns>
        public static byte[] ReadHttpResponse(HttpWebResponse response)
        {
            Stream respStream = response.GetResponseStream();
            List<byte> responseBytesList = new List<byte>();
            int read;
            do
            {
                read = respStream.ReadByte();
                if (read != -1)
                {
                    byte singleByte = (byte)read;
                    responseBytesList.Add(singleByte);
                }
            }
            while (read != -1);

            return responseBytesList.ToArray();
        }

        /// <summary>
        /// The method to send NotificationWait request to the server.
        /// </summary>        
        /// <param name="requestBody">The NotificationWait request body.</param>
        /// <returns>Return the NotificationWait response body.</returns>
        public NotificationWaitSuccessResponseBody NotificationWaitCall(IRequestBody requestBody)
        {
            string requestType = "NotificationWait";
            HttpWebResponse response = SendMAPIHttpRequest(this.site, this.mailStoreUrl, this.userName, this.domain, this.userPassword, requestBody, requestType, this.cookies);

            NotificationWaitSuccessResponseBody result = null;

            string responseCode = response.Headers["X-ResponseCode"];

            byte[] rawBuffer = ReadHttpResponse(response);

            response.GetResponseStream().Close();

            if (int.Parse(responseCode) == 0)
            {
                ChunkedResponse chunkedResponse = ChunkedResponse.ParseChunkedResponse(rawBuffer);
                NotificationWaitSuccessResponseBody responseSuccess = NotificationWaitSuccessResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
                result = responseSuccess;
            }
            else
            {
                this.site.Assert.Fail("MAPIHTTP call failed, the error code returned from server is: {0}", responseCode);
            }

            return result;
        }

        /// <summary>
        /// The private method to connect the exchange server through MAPIHTTP transport.
        /// </summary>
        /// <param name="mailStoreUrl">The mail store Url.</param>
        /// <param name="domain">The domain the server is deployed.</param>
        /// <param name="userName">The domain account name</param>
        /// <param name="userDN">User's distinguished name (DN).</param>
        /// <param name="password">user Password.</param>
        /// <returns>If the method succeeds, the return value is 0. If the method fails, the return value is an implementation-specific error code.</returns>
        public uint Connect(string mailStoreUrl, string domain, string userName, string userDN, string password)
        {
            uint returnValue = 0;
            this.domain = domain;
            this.userPassword = password;
            this.mailStoreUrl = mailStoreUrl;
            this.userName = userName;

            ConnectRequestBody connectBody = new ConnectRequestBody();
            connectBody.UserDN = userDN;

            // The server MUST NOT compress ROP response payload (rgbOut) or auxiliary payload (rgbAuxOut). 
            connectBody.Flags = 0x00000001;

            // The code page in which text data is sent.
            connectBody.Cpid = 1252;

            // The local ID for everything other than sorting.
            connectBody.LcidString = 0x00000409;

            // The local ID for sorting.
            connectBody.LcidSort = 0x00000409;
            connectBody.AuxiliaryBufferSize = 0;
            connectBody.AuxiliaryBuffer = new byte[] { };
            if (this.cookies == null)
            {
                this.cookies = new CookieCollection();
            }

            HttpWebResponse response = SendMAPIHttpRequest(this.site, mailStoreUrl, userName, domain, password, connectBody, "Connect", this.cookies);
            string transferEncoding = response.Headers["Transfer-Encoding"];
            string pendingInterval = response.Headers["X-PendingPeriod"];
            string responseCode = response.Headers["X-ResponseCode"];

            if (transferEncoding != null)
            {
                if (string.Compare(transferEncoding, "chunked") == 0)
                {
                    byte[] rawBuffer = ReadHttpResponse(response);

                    returnValue = uint.Parse(responseCode);
                    if (returnValue == 0)
                    {
                        ChunkedResponse chunkedResponse = ChunkedResponse.ParseChunkedResponse(rawBuffer);
                        ConnectSuccessResponseBody responseSuccess = ConnectSuccessResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
                    }
                    else
                    {
                        this.site.Assert.Fail("Can't connect the server through MAPI over HTTP, the error code is: {0}", responseCode);
                    }
                }
            }

            response.GetResponseStream().Close();
            this.cookies = response.Cookies;

            return returnValue;
        }

        /// <summary>
        /// Send ROP request through MAPI over HTTP
        /// </summary>
        /// <param name="rgbIn">ROP request buffer.</param>
        /// <param name="pcbOut">The maximum size of the rgbOut buffer to place response in.</param>
        /// <param name="rawData">The response payload bytes.</param>
        /// <returns>0 indicates success, other values indicate failure. </returns>
        public uint Execute(byte[] rgbIn, uint pcbOut, out byte[] rawData)
        {
            uint ret = 0;

            ExecuteRequestBody executeBody = new ExecuteRequestBody();

            // Client requests server to not compress or XOR payload of rgbOut and rgbAuxOut.
            executeBody.Flags = 0x00000003;
            executeBody.RopBufferSize = (uint)rgbIn.Length;
            executeBody.RopBuffer = rgbIn;

            // Set the max size of the rgbAuxOut
            executeBody.MaxRopOut = pcbOut; // 0x10008;
            executeBody.AuxiliaryBufferSize = 0;
            executeBody.AuxiliaryBuffer = new byte[] { };

            HttpWebResponse response = SendMAPIHttpRequest(
                this.site,
                this.mailStoreUrl,
                this.userName,
                this.domain,
                this.userPassword,
                executeBody,
                "Execute",
                this.cookies);

            string responseCode = response.Headers["X-ResponseCode"];

            byte[] rawBuffer = ReadHttpResponse(response);
            response.GetResponseStream().Close();

            if (int.Parse(responseCode) == 0)
            {
                ChunkedResponse chunkedResponse = ChunkedResponse.ParseChunkedResponse(rawBuffer);

                ExecuteSuccessResponseBody responseSuccess = ExecuteSuccessResponseBody.Parse(chunkedResponse.ResponseBodyRawData);

                rawData = responseSuccess.RopBuffer;
                ret = responseSuccess.ErrorCode;
            }
            else
            {
                rawData = null;
                this.site.Assert.Fail("MAPIHTTP call failed, the error code returned from server is: {0}", responseCode);
            }

            this.cookies = response.Cookies;
            return ret;
        }

        /// <summary>
        /// This method sends the disconnect request through MAPIHTTP transport to the server.
        /// </summary>
        /// <returns>If the method succeeds, the return value is 0. If the method fails, the return value is an implementation-specific error code.</returns>
        public uint Disconnect()
        {
            uint returnValue = 0;
            DisconnectRequestBody disconnectBody = new DisconnectRequestBody();
            disconnectBody.AuxiliaryBufferSize = 0;
            disconnectBody.AuxiliaryBuffer = new byte[] { };
            HttpWebResponse response = SendMAPIHttpRequest(this.site, this.mailStoreUrl, this.userName, this.domain, this.userPassword, disconnectBody, "Disconnect", this.cookies);
            string transferEncoding = response.Headers["Transfer-Encoding"];
            string pendingInterval = response.Headers["X-PendingPeriod"];
            string responseCode = response.Headers["X-ResponseCode"];

            if (transferEncoding != null)
            {
                if (string.Compare(transferEncoding, "chunked") == 0)
                {
                    byte[] rawBuffer = ReadHttpResponse(response);

                    if (uint.Parse(responseCode) == 0)
                    {
                        ChunkedResponse chunkedResponse = ChunkedResponse.ParseChunkedResponse(rawBuffer);
                        DisconnectSuccessResponseBody responseSuccess = DisconnectSuccessResponseBody.Parse(chunkedResponse.ResponseBodyRawData);
                        returnValue = responseSuccess.ErrorCode;
                    }
                }
            }
            
            response.GetResponseStream().Close();
            this.cookies = null;

            return returnValue;
        }
    }
}