using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertFileUtility
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string filePath = @"D:\1.docx";
                string saveFile = @"D:\Test\";
                Console.WriteLine("Start ..." + DateTime.Now.ToString());
                ExportFileUtility.EditFile(saveFile);
                Console.WriteLine("end ..." + DateTime.Now.ToString());
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
           // Console.ReadLine();
        }
    }
}
