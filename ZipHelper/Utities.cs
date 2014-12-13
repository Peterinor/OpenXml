using System.Security.Cryptography;
using System.Text;

namespace ZipHelper
{
    public class Utities
    {
        public static string GetMd5(string sDataIn)
        {

            MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider();

            byte[] bytValue = Encoding.UTF8.GetBytes(sDataIn);

            byte[] bytHash = md5.ComputeHash(bytValue);

            md5.Clear();

            StringBuilder sTemp = new StringBuilder();

            foreach (byte t in bytHash)
            {
                sTemp.Append(t.ToString("X").PadLeft(2, '0'));
            }

            return sTemp.ToString().ToLower();

        }

    }
}
