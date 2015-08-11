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