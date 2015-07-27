//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------
namespace Microsoft.Protocols.TestSuites.MS_COPYS
{
    using System;
    using System.IO;

    /// <summary>
    /// A class contains helper methods for file URL.
    /// </summary>
    public class FileUrlHelper
    {
        /// <summary>
        /// A method used to validate the file URL and get the file name if the file URL is a valid file URL.
        /// </summary>
        /// <param name="fileUrl">A parameter represents the file URL which is to be validated.</param>
        /// <returns>A return value represents the file name which the file URL points to.</returns>
        public static string ValidateFileUrl(string fileUrl)
        {
            if (string.IsNullOrEmpty(fileUrl))
            {
                throw new ArgumentNullException("fileUrl");
            }

            Uri fileLocation;
            if (!Uri.TryCreate(fileUrl, UriKind.Absolute, out fileLocation))
            {
                throw new UriFormatException(string.Format(@"The file URL should be a valid absolute URL. Actual:[{0}]", fileUrl));
            }

            string fileName = Path.GetFileName(fileUrl);
            if (string.IsNullOrEmpty(fileName))
            {
                throw new UriFormatException(string.Format(@"The file URL should point to a file. Actual:[{0}]", fileUrl));
            }

            return fileName;
        }
    }
}