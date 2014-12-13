using System;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;

namespace ZipHelper
{
    public class UnZip
    {
        public void UnZipToDir(string zipfile, string dir)
        {
            ZipInputStream s = new ZipInputStream(File.OpenRead(zipfile));

            ZipEntry theEntry;
            while ((theEntry = s.GetNextEntry()) != null)
            {

                string directoryName = Path.GetDirectoryName(dir);
                string fileName = Path.GetFileName(theEntry.Name);

                //生成解压目录
                Directory.CreateDirectory(directoryName);

                if (fileName != String.Empty)
                {
                    string file = dir + theEntry.Name;
                    string path = Path.GetDirectoryName(file);
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                    //解压文件到指定的目录
                    FileStream streamWriter = File.Create(file);

                    int size = 2048;
                    byte[] data = new byte[2048];
                    while (true)
                    {
                        size = s.Read(data, 0, data.Length);
                        if (size > 0)
                        {
                            streamWriter.Write(data, 0, size);
                        }
                        else
                        {
                            break;
                        }
                    }

                    streamWriter.Close();
                }
            }
            s.Close();
        }
    }
}
