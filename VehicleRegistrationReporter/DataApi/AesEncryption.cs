using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace VehicleRegistrationReporter.DataApi
{
    public class AesEncryption
    {
        /// <summary>
        ///  AES 加密
        /// </summary>
        /// <param name="str">明文</param>
        /// <param name="aesKey">密钥</param>
        /// <returns></returns>
        public static string AesEncrypt(string str, string aesKey)
        {
            if (str == null || aesKey == null)
            {
                return null;
            }
            string data = string.Empty;
            if (!string.IsNullOrEmpty(str) && !string.IsNullOrEmpty(aesKey))
            {
                byte[] toEncryptArray = Encoding.UTF8.GetBytes(str);
                using (Aes aes = Aes.Create())
                {
                    var iv = new byte[16];
                    for (int i = 0; i < iv.Length; i++)
                        iv[i] = 0;
                    aes.IV = iv;
                    aes.Key = Encoding.UTF8.GetBytes(aesKey);
                    aes.Mode = CipherMode.ECB;
                    aes.Padding = PaddingMode.PKCS7;
                    aes.BlockSize = 128;
                    var cryptoTransform = aes.CreateEncryptor();
                    var resultArray = cryptoTransform.TransformFinalBlock(toEncryptArray, 0, toEncryptArray.Length);
                    data = Convert.ToBase64String(resultArray, 0, resultArray.Length);

                }
            }
            return data;
        }

        /// <summary>
        ///  AES 解密
        /// </summary>
        /// <param name="str">密文</param>
        /// <param name="aesKey">密钥</param>
        /// <returns></returns>
        public static string AesDecrypt(string str, string aesKey)
        {
            if (str == null || aesKey == null)
            {
                return null;
            }
            string data = string.Empty;
            if (!string.IsNullOrEmpty(str) && !string.IsNullOrEmpty(aesKey))
            {
                byte[] toEncryptArray = Convert.FromBase64String(str);
                using (Aes aes = Aes.Create())
                {
                    var iv = new byte[16];
                    for (int i = 0; i < iv.Length; i++)
                        iv[i] = 0;
                    aes.IV = iv;
                    aes.Key = Encoding.UTF8.GetBytes(aesKey);
                    aes.Mode = CipherMode.ECB;
                    aes.Padding = PaddingMode.PKCS7;
                    aes.BlockSize = 128;
                    var cryptoTransform = aes.CreateDecryptor();
                    byte[] resultArray = cryptoTransform.TransformFinalBlock(toEncryptArray, 0, toEncryptArray.Length);
                    data = Encoding.UTF8.GetString(resultArray);
                }

            }
            return data;
        }

    }
}
