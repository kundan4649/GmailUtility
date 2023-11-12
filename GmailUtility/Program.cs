using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using System.Data;
using System.Diagnostics;
using System.Net.Mail;
using System.Net;
using Microsoft.Data.SqlClient;

class Program
{
    static void Main(string[] args)
    {

        //Timer timer = new Timer(DoWork, null, 0, 55000);
        //Console.WriteLine("Scheduler started. Press Enter to exit.");
        //Console.ReadLine();

        
        while (true)
        {
            var messages = GetMessages(userId: "me");

            //IEnumerable<Google.Apis.Gmail.v1.Data.Message> filteredMessages = messages.Where(msg =>
            //{
            //    // Find the "Date" header in the payload's headers
            //    var dateHeader = msg.Payload.Headers.FirstOrDefault(header => header.Name == "Date");
            //
            //    if (dateHeader != null)
            //    {
            //        if (DateTime.TryParse(dateHeader.Value, out DateTime messageDate))
            //        {
            //            // Check if the message's date is within the last 10 minutes
            //            return messageDate > DateTime.Now.AddMinutes(-10);
            //        }
            //    }
            //
            //    return false; // Message doesn't have a valid "Date" header or it's not within the last 10 minutes
            //});

            foreach (var msg in messages)
            {
                var dateHeader = msg.Payload.Headers.FirstOrDefault(header => header.Name == "Date");
                string dateString = (dateHeader == null || dateHeader.Value == null) ? "" : dateHeader.Value;
                string format = "ddd, d MMM yyyy HH:mm:ss zzz";

                //DateTime dateTime = DateTime.ParseExact(dateString, format, System.Globalization.CultureInfo.InvariantCulture);
                DateTime dateTime = DateTime.ParseExact(dateString, format, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AssumeUniversal);


                if (dateTime < DateTime.Now.AddMinutes(-3))
                {
                    continue;
                }

                var subjectHeader = msg.Payload.Headers.FirstOrDefault(header => header.Name.ToLower() == "subject");
                string subjectString = (subjectHeader == null || subjectHeader.Value == null) ? "" : subjectHeader.Value.ToLower();
                //string scriptFile = "";

                if (subjectString == "nifty-future")
                    RunScript("niftyfuture");
                else if (subjectString == "nifty-only")
                    RunScript("nifty");
                else if (subjectString == "nifty-option")
                    RunScript("niftyoption");
                else if (subjectString == "nifty-bankoption")
                    RunScript("bankniftyoption");
                else if (subjectString == "nifty-stockoption")
                    RunScript("niftystockoption");

                if (subjectString.StartsWith("nifty"))
                {
                    SendEmail();
                    Console.WriteLine($"{subjectHeader.Name}: {subjectHeader.Value}");
                    break;
                }


            }
            Console.WriteLine($"Current Time: {DateTime.Now}");
            System.Threading.Thread.Sleep(180 * 1000);
        }
    }

    private static void RunScript(string scriptName)
    {
        //RunScript("niftyfuture");
        string fileName = "NiftyProcessor";
        switch (scriptName.ToLower())
        {
            case "nifty":
                fileName = "NiftyProcessor";
                break;
            case "niftyfuture":
                fileName = "NiftyFutureProcessor";
                break;
            case "niftyoption":
                fileName = "OptionChainProcessor";
                break;
            case "bankniftyoption":
                fileName = "OptionChainProcessorBankNifty";
                break;
            case "niftystockoption":
                fileName = "StockOptionChainProcessor";
                break;
            default:
                break;
        }
        Process nodeProcess = new Process();
        nodeProcess.StartInfo.FileName = "C:\\Program Files\\nodejs\\node.exe"; // path to the node executable
        nodeProcess.StartInfo.Arguments = $"C:\\Users\\dell\\Documents\\Puppeteer\\OptionChain\\{fileName}.js";
        nodeProcess.Start();
        nodeProcess.WaitForExit();



    }
    private static IEnumerable<Google.Apis.Gmail.v1.Data.Message>? GetMessages(string userId, string[] labels = null, bool includeSpamAndTrash = false)
    {
        var clientId = "";
        var secretKey = "";

        ClientSecrets secrets = new ClientSecrets()
        {
            ClientId = clientId,
            ClientSecret = secretKey
        };

        UserCredential credential = GoogleWebAuthorizationBroker.AuthorizeAsync(secrets, new[] {
            GmailService.Scope.GmailSend, GmailService.Scope.GmailReadonly, GmailService.Scope.GmailSettingsBasic
        }, user: "user", CancellationToken.None).Result;

        //GoogleWebAuthorizationBroker.ReauthorizeAsync(credential, CancellationToken.None);

        using (var gmailService = new GmailService(new BaseClientService.Initializer() { HttpClientInitializer = credential }))
        {
            var lstMessageRequest = gmailService.Users.Messages.List(userId);

            lstMessageRequest.Q = "from:kundan.singh312201@gmail.com subject:nifty";
            //lstMessageRequest.Q = "from:kundan.singh312201@gmail.com subject:Logs AT: after:" + DateTime.UtcNow.AddMinutes(-3).ToString("yyyy/MM/dd HH:mm:ss");
            lstMessageRequest.MaxResults = 5; // Set the maximum number of messages to retrieve

            lstMessageRequest.IncludeSpamTrash = includeSpamAndTrash;
            if (labels != null)
                lstMessageRequest.LabelIds = labels;
            //bool hasNext = true;

            //while (hasNext) { 
            var messages = lstMessageRequest.Execute();
            //hasNext = messages.NextPageToken != null;
            if (messages.Messages != null)
                foreach (var message in messages.Messages)
                {
                    yield return GetSingleMessage(userId, message.Id, gmailService);
                    //if (hasNext)
                    //  lstMessageRequest.PageToken = messages.NextPageToken;
                }
            //}

        }

        //   return null;
    }

    private static Google.Apis.Gmail.v1.Data.Message GetSingleMessage(string userId, string messageId, GmailService service)
    {
        var getSingleMessageRequest = service.Users.Messages.Get(userId, messageId);
        getSingleMessageRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Metadata;
        getSingleMessageRequest.MetadataHeaders = new[] { "From", "To", "Subject", "Date" };

        return getSingleMessageRequest.Execute();
    }
    private static void SendEmail()
    {
        string value = GetLogData();
        if (string.IsNullOrEmpty(value))
            value = "No data found";


        var clientId = "";
        var secretKey = "";

        ClientSecrets secrets = new ClientSecrets()
        {
            ClientId = clientId,
            ClientSecret = secretKey
        };

        UserCredential credential = GoogleWebAuthorizationBroker.AuthorizeAsync(secrets, new[] {
            GmailService.Scope.GmailSend, GmailService.Scope.GmailReadonly, GmailService.Scope.GmailSettingsBasic
        }, user: "user", CancellationToken.None).Result;
        using (var gmailService = new GmailService(new BaseClientService.Initializer() { HttpClientInitializer = credential }))
        {
            //var profile =  gmailService.Users.GetProfile("userId:me").Execute();
            //Console.WriteLine(profile.EmailAddress);
            //Console.ReadLine();

            var msg = new Google.Apis.Gmail.v1.Data.Message
            {
                Raw = Base64UrlEncode("From: kundan.singh312201@gmail.com\r\n" +
                                   "To: kundan.singh312201@gmail.com\r\n" +
                                   $"Subject: Logs AT: {DateTime.Now}\r\n" +
                                   "Content-Type: text/plain; charset=utf-8\r\n\r\n" +
                                   value)
            };

            try
            {


                var request = gmailService.Users.Messages.Send(msg, "me");
                var response = request.Execute();
                Console.WriteLine("Email sent successfully!");

                //var request = service.Users.Messages.Send(message, "kundan.singh312201@gmail.com");
                //var response = request.Execute();
                //Console.WriteLine("Email sent successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }

        }
    }

    private static string GetLogData()
    {
        string value = string.Empty;
        string connectionString = "Data Source=DESKTOP-1D62623;Initial Catalog=NSE;User ID=sa;Password=;Connect Timeout=30;Encrypt=False;Trust Server Certificate=False;Application Intent=ReadWrite;Multi Subnet Failover=False";
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();

            //string sqlQuery = "SELECT * FROM APILogs  WHERE ExecutionDate > DATEADD(MINUTE,-5, GETDate()) AND APIStatus = '200'";
            string sqlQuery = "SELECT top 4 * FROM APILogs  WHERE APIStatus = '200'and APIName <> 'NiftyStockOption'  order by ExecutionDate desc";
            using (SqlCommand command = new SqlCommand(sqlQuery, connection))
            {
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                foreach (DataRow row in dataTable.Rows)
                {
                    // Add more columns as needed
                    // ExecutionStatus: {row["ExecutionStatus"]}, ExecutionDate: {row["ExecutionDate"]},
                    value += $"\r\n\t APIName: {row["APIName"]}, APIStatus: {row["APIStatus"]},  [Message]: {row["Message"]}, AffectedRecords: {row["AffectedRecords"]}, ExecutionDate: {row["ExecutionDate"]}";
                }
            }
        }
        return value;
    }

    private static string Base64UrlEncode(string input)
    {
        var inputBytes = System.Text.Encoding.UTF8.GetBytes(input);
        return Convert.ToBase64String(inputBytes)
            .Replace('+', '-')
            .Replace('/', '_')
            .TrimEnd('=');
    }

    private static void DoWork(object state)
    {
        List<string> lstTime = new List<string>() { "09:22", "09:33", "10:03", "10:33", "11:03", "11:33", "12:03", "12:33", "13:03", "13:33", "14:03", "14:33", "15:03", "15:33", "16:03" };
        //if (lstTime.Contains($"{DateTime.Now.Hour:00}:{DateTime.Now.Minute:00}"))
            SendEmail();
        // Put your scheduled task logic here
        Console.WriteLine($"Task executed at: {DateTime.Now}");
    }
}
