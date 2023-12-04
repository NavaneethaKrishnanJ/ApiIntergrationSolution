using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace TestExcel
{
    public static class HttpClientClass
    {
        public static async Task SendDataAsync(RequestModel excelData)
        {

            using (var clientHandler = new HttpClientHandler())
            {

                //ServicePointManager.ServerCertificateValidationCallback = (sender, cert, chain, sslPolicyErrors) => true;
                //clientHandler.ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => { return true; };
                //var client = new HttpClient(clientHandler);



                //var content = new ByteArrayContent(excelData.AttendanceWSFile);
                //client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", "Q0hBVkFSQVdTMTAwMDEzOmNoYXZhcmFsb2dpbg==");
                //client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
                //client.DefaultRequestHeaders.Add("Unit", "CHAVARA");
                //client.DefaultRequestHeaders.Add("ModuleType", "Attendance");
                //client.DefaultRequestHeaders.Add("RequestId", "CHAVARAATReq00013");
                //client.DefaultRequestHeaders.Add("AttendanceDate", shortDate);

                //HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, APIURL);
                //request.Content = content;
                //request.Headers.Add("AttendanceWSFile", myFileName);


                //await client.SendAsync(request)
                //      .ContinueWith(responseTask =>
                //      {
                //          var result = responseTask.Result;
                //          Console.WriteLine("Response: {0}", result.StatusCode);
                //          Console.WriteLine("Response: {0}", result.RequestMessage);
                //          Console.WriteLine("Response: {0}", result.IsSuccessStatusCode);
                //      });

            }
        }
    }
}
