using RestSharp;
using System.Configuration;
using System.Net;
using System.Timers;

namespace TestExcel
{
    public static class RestHttpClient
    {
        
        public static async Task sendRequestAsync(string myFileName, string Date )
        {
            string APIURL = ConfigurationSettings.AppSettings["APIURL"];
            string endPoint = ConfigurationSettings.AppSettings["Endpoint"];
            string AuthKey = ConfigurationSettings.AppSettings["AuthorizationKey"];
            string UnitName = ConfigurationSettings.AppSettings["UnitName"];
            string ModuleType = ConfigurationSettings.AppSettings["ModuleType"];
            string RequestId = ConfigurationSettings.AppSettings["RequestId"];
            string Fname = ConfigurationSettings.AppSettings["FileName"] != null ? ConfigurationSettings.AppSettings["FileName"] : myFileName;
            Fname = string.IsNullOrEmpty(Fname) ? myFileName : Fname;

            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            
            var options = new RestClientOptions(APIURL)
            {
                MaxTimeout = -1,
            };

            options.RemoteCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true; 
            var client = new RestClient(options);
            var request = new RestRequest(endPoint, Method.Post);
            request.AddHeader("Authorization", AuthKey);
            request.AddHeader("Cookie", "JSESSIONID=24362382630DE24E8A26230C16399FD3.jvm1; OFBiz.Visitor=3180018");
            request.AlwaysMultipartFormData = true;
            request.AddFile("AttendanceWSFile", Fname);
            request.AddParameter("Unit", UnitName);
            request.AddParameter("ModuleType", ModuleType);
            request.AddParameter("RequestId", RequestId);
            request.AddParameter("AttendanceDate", DateTime.Now);

            await client.ExecuteAsync(request).ContinueWith(x =>
            {
                Console.Write("Loading");
                for (int i = 0; i < 10; i++)
                {
                    Thread.Sleep(1000);
                    Console.Write(".");
                }
                var result = x.Result;
                Console.WriteLine(result?.Content?.ToString());
            });

        }
    }
}
