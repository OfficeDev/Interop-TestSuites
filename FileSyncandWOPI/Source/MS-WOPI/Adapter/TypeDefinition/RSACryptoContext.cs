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
    using System.Security.Cryptography;

    /// <summary>
    /// A class is used to support RSA crypt operation.
    /// </summary>
    public static class RSACryptoContext
    {  
        /// <summary>
        /// A string represents the unique key-pairs which is match the Asymmetric encrypt algorithm. It is used for "X-WOPI-Proof" header.
        /// </summary>
        private const string AsymmetricEncryptKeypairsOfCurrent = @"BwIAAACkAABSU0EyAAQAAAEAAQBxwXpxCIYvyvtnFmflVBmFEYpn/hhuZCqVH1PhqnAQr/ONAVkiONeMTToP7n2kUi1wntw5MbMaoIPWoNejZOLDgIVUqfjCOH2EbXOMmp6zTN35zAYbGZ3XWgLtmVkHIU60IQGvl7rOEHnEJ4v7a7Q2s6r4IOdVqFMS1T2YpmT6r5W6lKKyvjRKtwu3RClOcoNR9cpQNlzRqP6Tsl1B2UHAlRhNOBB7jEcttjdFFr/C1M6e5+XpDIhDlJt4BMODGN1tEAYZ71eMpnXkUBpUewXyaGZSSi5H/1cSki0srmONVNj7amPdk8QdEnlz+WnhLSjjeyNBHCVhhtVYaKfxd8LLnfRqGmPxGkcLsbTTL9Ngv1fNBFGCq45NsSNq4MGp+eja+2oJSE6duSpjeSOapLQ/vPcfkVZQP3AmOvEvxKV3wHUnzIlbvaklNhbg2LAJOZ2lT7bWVHfrgQ06lcjlapAaAoSNbPBhhhSOdUPrdRy4ebAhoUiriJiXoMNy9NS6GHqploWkCU/HDdVBTYS/yjqVFOAbhQA4edSgzyIT5P1tyIWImQ7ziE+7gFMPbXosoWmDL/iSmaAyuSU3lqun3lrBIWTXuHul48OgGadg2k2c6PBX0y7fqgBfEOAOFSii+c/d1G2umh321WigSaeg0VrsPO5pH1MbFOOFS/ZjMGWmZSYZfopkIOCqXL1UbiNwsIa0OG1FxrIMI9zgQt1aGPV4NK7unMeBz/t9uO2Y9nsrztQnYuje6K0cPTK+HlOExao=";
        
        /// <summary>
        /// A string represents the public key part of the unique key-pairs which is represent in "AsymmetricEncryptKeypairsOfCurrent" field.
        /// </summary>
        private const string PublicKeyValueOfCurrent = @"BgIAAACkAABSU0ExAAQAAAEAAQBxwXpxCIYvyvtnFmflVBmFEYpn/hhuZCqVH1PhqnAQr/ONAVkiONeMTToP7n2kUi1wntw5MbMaoIPWoNejZOLDgIVUqfjCOH2EbXOMmp6zTN35zAYbGZ3XWgLtmVkHIU60IQGvl7rOEHnEJ4v7a7Q2s6r4IOdVqFMS1T2YpmT6rw==";

        /// <summary>
        /// A string represents the unique key-pairs which is match the Asymmetric encrypt algorithm. It is used for "X-WOPI-ProofOld" header.
        /// </summary>
        private const string AsymmetricEncryptKeypairsOfOld = @"BwIAAACkAABSU0EyAAQAAAEAAQC3T8ExrB2fjcvpVJF7ZYbAh9yfsHsXMcqHa/0i0ncEdoejYr1s1NMbZtGbautAmDH2Q5/dUoZ6UHvymDxGh3VypfCHg7heRaPoeBLBrKyhIbG8oy2KUlpUSBGi9s2ZTb4tMyef8ZTA+f5jneAIZDC8U4DZF0mifHJtXrQHqSY9kkHv/7WdvxVsoLToq78tX3CZDR5btKGsOJD8qjwJ0Tthsq3l79rhh39kxci9YzMKVK4rQVlUSAopVtRuWXa2j8X3eOs/YEmObMpUxEPK6yZk8Rj9UYLMm7rmO+iB0vMrTAkKOI7csfDEg+XZKSM3tRmkOJHPyUlxgeLBcR7PDX+9gVPEILISnGGj+2qsHT+ywcmC7/dYiwjhj/VXgzZvl2KjDbfaa2N10CQN6MnBMXzawvsnrSA4x8UGmzLuzLiEuTlg6Ed1WuL7uv3p+Rs9NX/yuj2jPuuFWDNNWlqGb4455eERQv73YXLNFMGRc87peBib/gN2YUO5suH8J5NzyUOGA0YPEvNWHIZo0k0JcJWzY3zVLENdKxHZFjc60bfRbAM0wTl20aWEvUEUBPOikuRmeEFveJNYSGUvcscIefhtgdTzsafiOpUW/2nR+tpxoIQM4gFZXIs85838T46kNCX1M/RBLjailtbAuyjOhA8Dixjm2jL4nndpkKPRl5ZC/pmYnUZGVd/nbh7h9rKglRSLYFzc3+OnPI9Mj0UoDPy72SE4DO6e4QN7npbklrWAKSqGUF4DNT4Y7iw5pibqSw4=";

        /// <summary>
        /// A string represents the public key part of the unique key-pairs which is represent in "AsymmetricEncryptKeypairsOfOld" field.
        /// </summary>
        private const string PublicKeyValueOfOld = @"BgIAAACkAABSU0ExAAQAAAEAAQC3T8ExrB2fjcvpVJF7ZYbAh9yfsHsXMcqHa/0i0ncEdoejYr1s1NMbZtGbautAmDH2Q5/dUoZ6UHvymDxGh3VypfCHg7heRaPoeBLBrKyhIbG8oy2KUlpUSBGi9s2ZTb4tMyef8ZTA+f5jneAIZDC8U4DZF0mifHJtXrQHqSY9kg==";

        /// <summary>
        /// A binaries data represents the unique key-pairs which is represent in "AsymmetricEncryptKeypairsOfCurrent" field.
        /// </summary>
        private static byte[] fullkeysBlobOfCurrent = Convert.FromBase64String(AsymmetricEncryptKeypairsOfCurrent);

        /// <summary>
        /// A binaries data represents the unique key-pairs which is represent in "AsymmetricEncryptKeypairsOfOld" field.
        /// </summary>
        private static byte[] fullkeysBlobOfOld = Convert.FromBase64String(AsymmetricEncryptKeypairsOfOld);

        /// <summary>
        /// Gets the public key part of the unique key-pairs which is used for "X-WOPI-Proof" header.
        /// </summary>
        public static string PublicKeyStringOfCurrent
        {
            get 
            {
                return PublicKeyValueOfCurrent;
            }
        }

        /// <summary>
        /// Gets the public key part of the unique key-pairs which is used for "X-WOPI-ProofOld" header.
        /// </summary>
        public static string PublicKeyStringOfOld
        {
            get
            {
                return PublicKeyValueOfOld;
            }
        }

        /// <summary>
        /// Gets the binaries data for the public key part of the unique key-pairs which is used for "X-WOPI-Proof" header.
        /// </summary>
        public static byte[] PublicKeyBlobOfCurrent
        {
            get
            {
                return Convert.FromBase64String(PublicKeyValueOfCurrent);
            }
        }

        /// <summary>
        /// Gets the binaries data for the public key part of the unique key-pairs which is used for "X-WOPI-ProofOld" header.
        /// </summary>
        public static byte[] PublicKeyBlobOfOld
        {
            get
            {
                return Convert.FromBase64String(PublicKeyValueOfOld);
            }
        }

        /// <summary>
        /// A method is used to sign the data with old key-pairs. The signed data only pass the validation by using old public key.
        /// </summary>
        /// <param name="originalData">A parameter represents the binaries data which will be signed with old key-pairs.</param>
        /// <returns>A return value represents the signed data.</returns>
        public static byte[] SignDataWithOldPublicKey(byte[] originalData)
        {
            return SignData(originalData, fullkeysBlobOfOld);
        }

        /// <summary>
        /// A method is used to sign the data with current key-pairs. The signed data only pass the validation by using current public key.
        /// </summary>
        /// <param name="originalData">A parameter represents the binaries data which will be signed with current key-pairs.</param>
        /// <returns>A return value represents the signed data.</returns>
        public static byte[] SignDataWithCurrentPublicKey(byte[] originalData)
        {
            return SignData(originalData, fullkeysBlobOfCurrent);
        }

        /// <summary>
        /// A method is used to validate the signed data by using old public key.
        /// </summary>
        /// <param name="signedData">A parameter represents the signed data which will be validate.</param>
        /// <param name="originalData">A parameter represents the original data which is used to execute the validation.</param>
        /// <returns>Return 'true' indicating the signed data pass the validation.</returns>
        public static bool VerifySignedDataUsingOldKey(byte[] signedData, byte[] originalData)
        {
            return VerifySignedData(signedData, originalData, PublicKeyBlobOfOld);
        }

        /// <summary>
        /// A method is used to validate the signed data by using current public key.
        /// </summary>
        /// <param name="signedData">A parameter represents the signed data which will be validate.</param>
        /// <param name="originalData">A parameter represents the original data which is used to execute the validation.</param>
        /// <returns>Return 'true' indicating the signed data pass the validation.</returns>
        public static bool VerifySignedDataUsingCurrentKey(byte[] signedData, byte[] originalData)
        {
            return VerifySignedData(signedData, originalData, PublicKeyBlobOfCurrent);
        }

        /// <summary>
        /// A method is used to sign the data with specified key-pairs. The signed data only pass the validation by using the public key part of the specified key-pairs.
        /// </summary>
        /// <param name="originalData">A parameter represents the binaries data which will be signed with old key-pairs.</param>
        /// <param name="fullkeyBlob">A parameter represents the binaries data of the unique key-pairs which is match the asymmetric encrypt algorithm.</param>
        /// <returns>A return value represents the signed data.</returns>
        private static byte[] SignData(byte[] originalData, byte[] fullkeyBlob)
        {
            if (null == originalData || 0 == originalData.Length)
            {
                throw new ArgumentNullException("originalData");
            }

            if (null == fullkeyBlob || 0 == fullkeyBlob.Length)
            {
                throw new ArgumentNullException("fullkeyBlob");
            }

            using (RSACryptoServiceProvider rsaProvider = new RSACryptoServiceProvider())
            {
                rsaProvider.ImportCspBlob(fullkeyBlob);
                SHA256Managed sha = new SHA256Managed();
                byte[] signedData = rsaProvider.SignData(originalData, sha);
                sha.Dispose();
                return signedData;
            }
        }

        /// <summary>
        /// A method is used to validate the signed data by using specified public key.
        /// </summary>
        /// <param name="signedData">A parameter represents the signed data which will be validate.</param>
        /// <param name="originalData">A parameter represents the original data which is used to execute the validation.</param>
        /// <param name="publicKeyBlob">A parameter represents the binaries data of the public key part of a unique key-pairs.</param>
        /// <returns>Return 'true' indicating the signed data pass the validation.</returns>
        private static bool VerifySignedData(byte[] signedData, byte[] originalData, byte[] publicKeyBlob)
        {
            if (null == signedData || 0 == signedData.Length)
            {
                throw new ArgumentNullException("signedData");
            }

            if (null == originalData || 0 == originalData.Length)
            {
                throw new ArgumentNullException("originalData");
            }

            if (null == publicKeyBlob || 0 == publicKeyBlob.Length)
            {
                throw new ArgumentNullException("publicKeyBlob");
            }

            using (RSACryptoServiceProvider rsaProvider = new RSACryptoServiceProvider())
            {
                rsaProvider.ImportCspBlob(publicKeyBlob);
                SHA256Managed sha = new SHA256Managed();
                bool result = rsaProvider.VerifyData(originalData, sha, signedData);
                sha.Dispose();
                return result;
            }
        }
    }
}
