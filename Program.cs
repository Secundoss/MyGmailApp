using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using Google.Apis.Util.Store;


namespace MyGmailApp
{
    internal class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("Gmail API: START");
            Console.WriteLine("================================");
            try
            {
                new Program().Run().Wait();
            }
            catch (AggregateException ex)
            {
                foreach (var e in ex.InnerExceptions)
                {
                    Console.WriteLine("ERROR: " + e.Message);
                }
            }
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        private async Task Run()
        {
            UserCredential credential;
            string ApplicationName = "SupGmailWebHelper";
            string Path = Environment.CurrentDirectory;

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    new[] {GmailService.Scope.GmailReadonly}, // Add param "GmailService.Scope.GmailSend" to send messages
                "user", CancellationToken.None, new FileDataStore("MyGmailApp"));
            }

            // Create the service.
            var Service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Get labels dictionary.
            var LabelsList = Service.Users.Labels.List("me");
            var emailListResponse = LabelsList.Execute();

            Dictionary<string, string> LabelsDictionary = new Dictionary<string, string>();

            foreach (var labels in emailListResponse.Labels)
            {
                var emailInfoRequest = Service.Users.Labels.Get("me", labels.Id);
                var emailInfoResponse = emailInfoRequest.Execute();

                if (emailInfoResponse.Id.Contains("Label"))
                {
                    LabelsDictionary.Add(emailInfoResponse.Name, emailInfoResponse.Id);
                }
            }
            
            var LabelsNameList = LabelsDictionary.Keys.ToList();

            foreach (string Key in LabelsNameList)
            {
                Console.WriteLine(Key +":"+ GetThreadsCounts(Key, Service, Path).ToString());
            }
        }


        // Save data class.
        public void SaveDictionary(string filePath, Dictionary<string, string> labels)
        {
            using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate))
            {
                using (TextWriter tw = new StreamWriter(fs))
                    foreach (KeyValuePair<string, string> kvp in labels)
                    {
                        tw.WriteLine(string.Format("{0};{1}", kvp.Key, kvp.Value));
                    }
            }
        }

			  // Gmail Apis class.
        int GetThreadsCounts(string Label, GmailService Service, string Path)
        {
            var Request = Service.Users.Threads.List("me");
            Request.Q = $"is:unread / label:" + Label;
            var Response = Request.Execute();
            var Threads =  Response.Threads;
            return Threads.Count();
        }
    }
}
