using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.Data.SqlClient;

namespace Serivcemessgae
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                System.Threading.Timer timer = new System.Threading.Timer(timer_Elapsed, null, 0, 300000);

                Console.WriteLine("Press enter to end the program.");
                Console.ReadLine();
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc.Message + "  " + exc.InnerException);
            }

           

        }

        static void timer_Elapsed(object state)
        {

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            try
            {
              
                service.Url = new Uri("https://email.poloralphlauren.com/ews/exchange.asmx");

                service.Credentials = new NetworkCredential("gchauhan", "Ralph@1235", "PRLUS01");
                service.UseDefaultCredentials = false;

                FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, new ItemView(10));

                Item item = findResults.Items.OrderByDescending(x=>x.DateTimeReceived).First(x => x.IsFromMe);

                string sqlQuery = item.Subject.Substring(6);

                string inputValidation = item.Subject.Substring(0, 5);

                if (inputValidation == "query")
                {
                    string queryData = GetData(sqlQuery);

                    EmailMessage email = new EmailMessage(service);

                    email.ToRecipients.Add("gopal.chauhan@ralphlauren.com");

                    email.Subject = "QueryData";
                    email.Body = new MessageBody(queryData);

                    email.Send();
                }
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc.Message + "  " + exc.InnerException);
                if (exc.Message != "Sequence contains no matching element")
                {
                    EmailMessage email = new EmailMessage(service);

                    email.ToRecipients.Add("gopal.chauhan@ralphlauren.com");

                    email.Subject = "Exception";
                    email.Body = new MessageBody(exc.Message);

                    email.Send();
                }

                
            }
        }

        private static string GetData(string queryString)
        {
            string data = string.Empty;
            StringBuilder sb = new StringBuilder();
            string connectionString = "Data Source=USNYDVSQLDB2V;Initial Catalog=PMC;Trusted_Connection=yes;persist security info=False";         


            using (SqlConnection connection =
                new SqlConnection(connectionString))
            {

                SqlCommand command = new SqlCommand(queryString, connection);

                try
                {
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    sb.Append("<table><tbody>");
                    while (reader.Read())
                    {
                        sb.AppendFormat("<tr><td>{0}</td><td>{1}</td></tr>", reader[1], reader[2]);
                       
                    }
                    reader.Close();
                    sb.Append("</tbody></table>");
                }
                catch (Exception ex)
                {
                   data =  ex.Message;
                }

            }
            return data = sb.ToString();
        }
       
    }
}
