using Newtonsoft.Json;
using Org.BouncyCastle.Asn1.Ocsp;
using RestSharp;
using System.Configuration;
using System.Net;
using System.Net.Http.Headers;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace TestExcel
{
    internal class Program
    {
      

        public static DateTime myDate = DateTime.Now;
        public static string shortDate = myDate.ToString("yyyyMMdd");
        public static string SourceFolder = string.Empty;
        public static string DestinationFoler = string.Empty;
        public static string myFileName = string.Empty;
        static async Task  Main(string[] args)
        {

            // ExcelData data = new ExcelData();
            //data.GenerateExcel();



            SourceFolder = ConfigurationSettings.AppSettings["SourceFolder"];
            DestinationFoler = ConfigurationSettings.AppSettings["DestinationFoler"];

            DirectoryInfo d = new DirectoryInfo(SourceFolder); //Assuming Test is your Folder

            FileInfo[] Files = d.GetFiles("*.xlsx"); 
          
            string str = "";

            foreach (FileInfo file in Files)
            {
                str = file.Name;
                Console.WriteLine(str);
            }

            myFileName = SourceFolder + str;
            DestinationFoler =  DestinationFoler + str;

            if (File.Exists(myFileName))
            {
                Console.WriteLine("The file exists.");
                await RestHttpClient.sendRequestAsync(myFileName,shortDate);
                File.Move(myFileName, DestinationFoler);
                Console.ReadLine();
            }else
            {
                Console.WriteLine("The file does not exist.");
            }

        }

    }
}
