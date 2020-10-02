using Meziantou.Framework.Win32;
using Newtonsoft.Json;
using openapi_excel.Security;
using openapi_excel.UI;
using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Forms;

namespace openapi_excel
{
    public static class SecurityManager
    {
        static byte[] s_aditionalEntropy = { 9, 8, 7, 6, 5 };

        internal static void Login(bool relogin = false)
        {
            var appSettings = ConfigurationManager.AppSettings;
            var apiName = appSettings["ApiName"];
            var apiNameNoSpace = appSettings["ApiNameNoSpace"];
            var credSaveDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), apiNameNoSpace);

            if (!Directory.Exists(credSaveDirectory))
            {
                Directory.CreateDirectory(credSaveDirectory);
            }

            var credLocation = Path.Combine(credSaveDirectory, $"{apiNameNoSpace}.credentials");
            if (!relogin && File.Exists(credLocation))
            {
                var protectedData = File.ReadAllBytes(credLocation);
                var unprotectedJson = ProtectedData.Unprotect(protectedData, s_aditionalEntropy, DataProtectionScope.CurrentUser);

                var authCreds = JsonConvert.DeserializeObject<BasicAuthCredentials>(Encoding.UTF8.GetString(unprotectedJson));

                SwaggerRegistry.SaveCredentials(authCreds);
                return;
            }

            var creds = CredentialManager.PromptForCredentials(
                captionText: $"Login to {apiName}",
                messageText: $"Please login to {apiName}",
                saveCredential: CredentialSaveOption.Selected);

            if (creds != null)
            {
                var basicAuthCredentials = new BasicAuthCredentials { Username = creds.UserName, Password = creds.Password };
                var basicAuthAsJson = JsonConvert.SerializeObject(basicAuthCredentials);

                var protectedJson = ProtectedData.Protect(Encoding.UTF8.GetBytes(basicAuthAsJson), s_aditionalEntropy, DataProtectionScope.CurrentUser);
                File.WriteAllBytes(credLocation, protectedJson);
                SwaggerRegistry.SaveCredentials(basicAuthCredentials);
            }
        }

        internal static void Relogin()
        {
            Login(relogin: true);
        }

        public static void ShowApiSecurityForm()
        {
            if (SwaggerRegistry.Api == null)
            {
                System.Windows.MessageBox.Show("Api not loaded");
            }

            var appSettings = ConfigurationManager.AppSettings;
            var apiName = appSettings["ApiName"];

            var apiKeyWindow = new ApiKeys();
            apiKeyWindow.SetApiKeys(SwaggerRegistry.ApiKeyCredentials.Values.ToList());

            Window window = new Window
            {
                Title = $"{apiName} API Security",
                Content = apiKeyWindow,
                Height = 500,
                Width = 300,
                MinHeight = 250,
                MinWidth = 300,
                SizeToContent = SizeToContent.Height
            };

            window.ShowDialog();

            foreach (var apiKey in apiKeyWindow.apiThings)
            {
                SwaggerRegistry.ApiKeyCredentials[apiKey.Key] = new ApiKey { Key = apiKey.Key, Value = apiKey.Value, In = apiKey.In };
            }
        }
    }
}
