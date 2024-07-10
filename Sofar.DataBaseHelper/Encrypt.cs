﻿using System;
using System.Security.Cryptography;
using System.Text;

namespace Sofar.DataBaseHelper
{
    public class Encrypt
    {
        private const string Cryptography_Key = "ZTE_E-University_EmpTrain_DB_ConnString_Key_2006";

        private static string Decrypt3DES(string strValue, string strKey)
        {
            TripleDESCryptoServiceProvider provider = new TripleDESCryptoServiceProvider();
            MD5CryptoServiceProvider provider2 = new MD5CryptoServiceProvider();
            if (strValue == null)
            {
                return string.Empty;
            }
            provider.Key = provider2.ComputeHash(Encoding.ASCII.GetBytes(strKey));
            provider.Mode = CipherMode.ECB;
            ICryptoTransform transform = provider.CreateDecryptor();
            byte[] inputBuffer = Convert.FromBase64String(strValue);
            return Encoding.ASCII.GetString(transform.TransformFinalBlock(inputBuffer, 0, inputBuffer.Length));
        }

        private static string Decrypt3DES(string strValue, string strKey, Encoding encoding)
        {
            TripleDESCryptoServiceProvider provider = new TripleDESCryptoServiceProvider();
            provider.Key = new MD5CryptoServiceProvider().ComputeHash(encoding.GetBytes(strKey));
            provider.Mode = CipherMode.ECB;
            ICryptoTransform transform = provider.CreateDecryptor();
            string str = "";
            try
            {
                byte[] inputBuffer = Convert.FromBase64String(strValue);
                str = encoding.GetString(transform.TransformFinalBlock(inputBuffer, 0, inputBuffer.Length));
            }
            catch (Exception exception)
            {
                Console.WriteLine("错误：{0}", exception);
            }
            return str;
        }

        public static string DecryptPassword(string strPassword)
        {
            return Decrypt3DES(strPassword, "ZTE_E-University_EmpTrain_DB_ConnString_Key_2006");
        }

        private static string Encrypt3DES(string a_strString, string a_strKey)
        {
            TripleDESCryptoServiceProvider provider = new TripleDESCryptoServiceProvider();
            provider.Key = new MD5CryptoServiceProvider().ComputeHash(Encoding.ASCII.GetBytes(a_strKey));
            provider.Mode = CipherMode.ECB;
            ICryptoTransform transform = provider.CreateEncryptor();
            byte[] bytes = Encoding.ASCII.GetBytes(a_strString);
            return Convert.ToBase64String(transform.TransformFinalBlock(bytes, 0, bytes.Length));
        }

        private static string Encrypt3DES(string strValue, string strKey, Encoding encoding)
        {
            TripleDESCryptoServiceProvider provider = new TripleDESCryptoServiceProvider();
            provider.Key = new MD5CryptoServiceProvider().ComputeHash(encoding.GetBytes(strKey));
            provider.Mode = CipherMode.ECB;
            ICryptoTransform transform = provider.CreateEncryptor();
            byte[] bytes = encoding.GetBytes(strValue);
            return Convert.ToBase64String(transform.TransformFinalBlock(bytes, 0, bytes.Length));
        }

        public static string EncryptPassword(string strPassword)
        {
            return Encrypt3DES(strPassword, "ZTE_E-University_EmpTrain_DB_ConnString_Key_2006");
        }

        //public static string GetConString(string sComponentID)
        //{
        //    string str = string.Empty;
        //    IClientInfo clientInfoInstance = ConfigHandler.GetClientInfoInstance();
        //    if (clientInfoInstance != null)
        //    {
        //        str = clientInfoInstance.GetCurrentOrgID() + "_";
        //    }
        //    string strValue = ConfigHandler.GetDBLinkSettings(new string[] { str + sComponentID.ToUpper() })[0];
        //    return Decrypt3DES(strValue, "ZTE_E-University_EmpTrain_DB_ConnString_Key_2006");
        //}
    }
}

