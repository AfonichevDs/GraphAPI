using Newtonsoft.Json;
using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace OneNoteManager
{
    class Program
    {
        private const string tokensPath = @"C:\Temp\Tokens.json";
        private const string htmlPath = @"C:\Temp\template.html";
        private const string filePath = @"C:\Temp\fileCyb.pdf";

        private const string clientId = "{clientIdHere}";
        private const string aadInstance = "https://login.microsoftonline.com/{0}";
        private const string tenant = "charpcompany.onmicrosoft.com";
        private const string appKey = "{appKeyHere};
        private const string redirectUrl = "https://charpcompany.onmicrosoft.com/CompanyManager";
        private const string code = "OAQABAAIAAACQN9QBRU3jT6bcBQLZNUj76WDlKIZQtznfNDpwZVS7auzv13B2QcN9KLv1x5o2mb7-HRw6c00JLeGEbm3UBeaiKFJCAvUYL6R6O14tgNXJ0w_gkxRyIQgsO8Q5Pw9aW_up1iT7176RhJ0XbUIrgCvJ_XgWZXoq251bKfWod8anxN3dSI0zCujkQu9nkTWp_IwSryW91yceuDAuOf_CVqy_JzpQHwunEh9koHHEIqt9yByrVctN6Sc5NHSXvO77rBGA6K-b8-qoFWalZI0adscVk2KFkYFZIfyyzdRJH6CYQXr9liaTa3nDrimgPleMenWCHXABKk-hZHcd0tQRkrBFMs_78GfpyOdXyxYAycLiZYbV3Se-Q4N6_LEC-JaKIPQSHy1H-kkW-CovVzyDrUEG-7s62wnIA88we0sIxSAv_p1liEL_eRWAhWXOh5O35Hsv4B_lH9oTgg5sR9WfFv5k09FGlXiY8wq52ZmQp-noNB36nJkr8c6noeR5Yw0Msdcoe877qhYyC9F_DdHIImFpzPY9puxn9NH19Epbs4suBmrXah7066kW-yZCkXAk-saAViPfQagSPHI17xmJgJ0ASSZ8FZq_XanBY1s8t5Dmyi9-TClfDighBA7MbzsqEWqaASfwr8Yv1v_GKsbjQ8iC6gLjUtxMwyb5Gb3HS0kiIKsnSEjsYKT6-TymQ7qy8UBiTD-2FixjcKHNpKtE6Ikw_AXOqLZsC_oOgbC5PNUIbzYefYEHf8kRlZy9fjfj3hnNKLi0zHAGrMv4xFGbpKxHnhVZ__ULs4vKIddOKEGV96BancKVv59NgRb_JTOsu7MgAA";

        private static string baseLoginUrl = string.Format(aadInstance, tenant);  

        private static HttpClient httpClient = new HttpClient();

        private static string scopes = "offline_access%20Notes.Create%20Notes.Read%20Notes.Read.All%20Notes.ReadWrite%20Notes.ReadWrite.All%20User.Export.All%20User.Read%20User.Read.All%20User.ReadBasic.All%20User.ReadWrite%20User.ReadWrite.All%20email%20openid%20profile";
        private static string notesId = "1-f563209d-6f74-41e2-855f-a7769b2c9085"; //id note from oneNote
        private static string meID = "38ae8312-b96b-4385-adec-1a4f5207396c"; //id користувача з запиту GetMe
        private static string sectionId = "1-e3b1dd00-d773-4358-89cb-398e8a3937f5";
        private static string pageId = "1-2f2408b94bce429d900f6e3ec70e487a!84-e3b1dd00-d773-4358-89cb-398e8a3937f5";

        private static Tokens tokens;
        public static Tokens Tokens
        {
            get
            {
                if(tokens == null)
                    tokens = JsonConvert.DeserializeObject<Tokens>(File.ReadAllText(tokensPath));
                return tokens;
            }
            set
            {
                SaveTokens(value);
                tokens = null;
            }
        }

        static void Main(string[] args)
        {
            Task task = GetTokens();
            task.Wait();

            //Task newTask = GetAccessTokenByRefresh();
            //newTask.Wait();

            //Task<string> userTask = GetMe();
            //userTask.Wait();

            Task<string> notesTask = GetNotes();
            notesTask.Wait();

            Task pageTask = SendFile();
            pageTask.Wait();

            Console.WriteLine("Hello World!");
            Console.ReadKey();
        }


        public static async Task GetTokens()
        {
            StringContent content = new StringContent(@$"client_id={clientId}&scope={scopes}
                                                        &code={code} 
                                                        &redirect_uri=https%3A%2F%2Fcharpcompany.onmicrosoft.com%2FCompanyManager
                                                        &grant_type=authorization_code
                                                        &client_secret={appKey}", Encoding.UTF8, "application/x-www-form-urlencoded");
            using (var response = await httpClient.PostAsync($"{baseLoginUrl}/oauth2/v2.0/token", content))
            {
                if(response.IsSuccessStatusCode)
                {
                    string result = await response.Content.ReadAsStringAsync();
                    dynamic res = JsonConvert.DeserializeObject(result);
                    var tokens = new Tokens
                    {
                        AccessToken = res.access_token.ToString(),
                        RefreshToken = res.refresh_token.ToString()
                    };
                    Tokens = tokens;
                }
            }
        }

        public static async Task GetAccessTokenByRefresh()
        {
            StringContent content = new StringContent(@$"client_id={clientId}&scope={scopes}
                                                         &refresh_token={Tokens.RefreshToken}
                                                         &redirect_uri=https%3A%2F%2Fcharpcompany.onmicrosoft.com%2FCompanyManager
                                                         &grant_type=refresh_token
                                                         &client_secret={appKey}", Encoding.UTF8, "application/x-www-form-urlencoded");
            using (var response = await httpClient.PostAsync($"{baseLoginUrl}/oauth2/v2.0/token", content))
            {
                if(response.IsSuccessStatusCode)
                {
                    string result = await response.Content.ReadAsStringAsync();
                    dynamic res = JsonConvert.DeserializeObject(result);
                    var tokens = new Tokens
                    {
                        AccessToken = res.access_token.ToString(),
                        RefreshToken = res.refresh_token.ToString()
                    };
                    Tokens = tokens;
                }
            }
        }

        public static async Task<string> GetNotes()
        {
            httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", Tokens.AccessToken);
            using (var response = await httpClient.GetAsync($"https://graph.microsoft.com/v1.0/me/onenote/notebooks/{notesId}/sections"))
            {
                if(response.IsSuccessStatusCode)
                {
                    string result = await response.Content.ReadAsStringAsync();
                    return result;
                }
            }
            return null;
        }

        public static async Task GetPages()
        {
            httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", Tokens.AccessToken);
            using (var response = await httpClient.GetAsync($"https://graph.microsoft.com/v1.0/me/onenote/pages"))
            {
                if(response.IsSuccessStatusCode)
                {
                    string result = await response.Content.ReadAsStringAsync();
                }
            }
        }

        public static async Task<string> CreateNotes()
        {
            var htmlMessageBody = File.ReadAllText(htmlPath);
            StringContent content = new StringContent(htmlMessageBody, Encoding.UTF8, "application/xhtml+xml");
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", Tokens.AccessToken);

            using (var response = await httpClient.PostAsync($"https://graph.microsoft.com/v1.0/me/onenote/sections/{sectionId}/pages", content))
            {
                if(response.IsSuccessStatusCode)
                {
                    string result = await response.Content.ReadAsStringAsync();
                }
            }
            return null;
        }

        public static async Task SendFile()
        {
            var htmlMessageBody = File.ReadAllText(htmlPath);
            HttpContent content = new StringContent(htmlMessageBody, Encoding.UTF8, "application/xhtml+xml");
            HttpContent fileStreamContent = new StreamContent(File.OpenRead(filePath));

            httpClient.DefaultRequestHeaders.Remove("ContentType");
            httpClient.DefaultRequestHeaders.TryAddWithoutValidation("ContentType", "multipart/form-data; boundary=MyContentID");
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", Tokens.AccessToken);

            fileStreamContent.Headers.ContentDisposition = new ContentDispositionHeaderValue("form-data");
            fileStreamContent.Headers.ContentDisposition.Name = "fileBlock";
            fileStreamContent.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");

            content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("form-data");
            content.Headers.ContentDisposition.Name = "Presentation";

            using (var formData = new MultipartFormDataContent("MyContentID"))
            {
                formData.Add(content, "Presentation");
                formData.Add(fileStreamContent, "fileBlock", "cyb.pdf");
                using (var response = await httpClient.PostAsync($"https://graph.microsoft.com/v1.0/me/onenote/sections/{sectionId}/pages", formData))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        string result = await response.Content.ReadAsStringAsync();
                    }
                }
            }
        }

        public static async Task<string> GetMe()
        {
            httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", Tokens.AccessToken);
            using (var response = await httpClient.GetAsync($"https://graph.microsoft.com/v1.0/me/"))
            {
                if (response.IsSuccessStatusCode)
                {
                    string result = await response.Content.ReadAsStringAsync();
                    return result;
                }
            }
            return null;
        }

        public static async Task<string> CreateSectionNotes()
        {
            var data = "{\"displayName\":\"Section name\"}";

            StringContent content = new StringContent(data, Encoding.UTF8, "application/json");
            httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", Tokens.AccessToken);
            using (var response = await httpClient.PostAsync($"https://graph.microsoft.com/v1.0/me/onenote/notebooks/{notesId}/sections", content))
            {
                if (response.IsSuccessStatusCode)
                {
                    string result = await response.Content.ReadAsStringAsync();
                }
            }
            return null;
        }

        public static async Task GetNotesContent()
        {
            httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", Tokens.AccessToken);
            using (var response = await httpClient.GetAsync($"https://graph.microsoft.com/v1.0/me/onenote/pages/{pageId}/content?includeIDs=true"))
            {
                if(response.IsSuccessStatusCode)
                {
                    string result = await response.Content.ReadAsStringAsync();
                }
            }
        }


        public static void SaveTokens(Tokens tokens)
        {
            File.WriteAllText(tokensPath, JsonConvert.SerializeObject(tokens));
        }   
        
    }
}
