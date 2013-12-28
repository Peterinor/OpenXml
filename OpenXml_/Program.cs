using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ZipHelper;

using System.IO;

namespace OpenXml_
{
	class Program
	{
		static void Main(string[] args)
		{
            UnZip un = new UnZip();
            un.UnZipToDir("./word.docx", "./word/");


            Console.WriteLine(Path.Combine(new string[] { "./asf/", "/asd.psd" }));

            Console.WriteLine(Utities.GetMD5("dfasfd"));
		}
	}
}
