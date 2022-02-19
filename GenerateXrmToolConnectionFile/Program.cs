using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk.Discovery;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.ServiceModel.Description;

namespace GenerateXrmToolConnectionFile
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var username = "user.name@example.invalid";
            var password = ""; // not needed for online connections

            var clientId = "51f81489-12ee-4a9e-aaae-a2591f45987d";
            var replyUri = new Uri("app://58145B91-0C36-4500-8554-080854F2AC97");
            var discoveryServiceUrl = new Uri("https://globaldisco.crm.dynamics.com/api/discovery/v2.0/Instances");

            var connectionFileName = username;
            var connectionDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"MscrmTools\XrmToolBox\Connections");
            var connectionFilePath = Path.Combine(connectionDirectory, connectionFileName + ".xml");

            var credentials = new ClientCredentials();
            credentials.UserName.UserName = username;
            credentials.UserName.Password = password;

            /* There seems to be some sort of bug in LinqPad where it gets MethodMissing for
             * DiscoverGlobalOrganizations even though it can find it intellisense, and the
             * same code works in a console app. Bizarrely reflection does work. When moved
             * into a real VS solution, the commented out connection code is the way to go (or
             * when the LP bug is fixed)


            var organizations = CrmServiceClient.DiscoverGlobalOrganizations(discoveryServiceUri: discoveryServiceUrl,
                                                                             clientCredentials: credentials,
                                                                             user: null,
                                                                             clientId: appId,
                                                                             redirectUri: appReplyUri,
                                                                             // if you change the prompt behaviour to auto, you'll want to use a different cache for each credential
                                                                             // otherwise it will keep reusing the last successful connection
                                                                             tokenCachePath: "",
                                                                             isOnPrem: false,
                                                                             authority: string.Empty,
                                                                             promptBehavior: Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior.SelectAccount,
                                                                             useDefaultCreds: false);
            */

            var organizations = (OrganizationDetailCollection)typeof(CrmServiceClient)
                                    .GetMethod(nameof(CrmServiceClient.DiscoverGlobalOrganizations), BindingFlags.Static | BindingFlags.Public)
                                    .Invoke(null, new object[] {
                              discoveryServiceUrl,
                              credentials,
                              null,
                              clientId,
                              replyUri,
                              "",
                              false,
                              string.Empty,
                              Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior.SelectAccount,
                              false,
                                    });

            var connections = organizations.Select(o =>
            {
                var url = new Uri(o.Endpoints[EndpointType.WebApplication]);

                var detail = new ConnectionDetail(createNewId: true)
                {
                    ConnectionId = o.OrganizationId,
                    UserName = username,
                    UserDomain = credentials.Windows.ClientCredential.Domain ?? "",
                    AuthType = GuessAuthType(credentials, clientId, o),
                    NewAuthType = GuessNewAuthType(credentials, clientId, o),
                    ConnectionName = $"{o.FriendlyName}",
                    ServerName = url.Host,
                    Organization = o.UniqueName,
                    OrganizationUrlName = o.UrlName,
                    OrganizationFriendlyName = o.FriendlyName,
                    OrganizationVersion = o.OrganizationVersion,
                    OrganizationServiceUrl = o.Endpoints[EndpointType.OrganizationService],
                    OrganizationDataServiceUrl = o.Endpoints[EndpointType.OrganizationDataService],
                    WebApplicationUrl = o.Endpoints[EndpointType.WebApplication],
                    ServerPort = url.Port,
                    Timeout = TimeSpan.FromMinutes(20),
                    EnvironmentId = o.EnvironmentId,
                    OriginalUrl = url.AbsoluteUri,
                    ReplyUrl = replyUri.AbsoluteUri,
                };

                if (detail.NewAuthType == AuthenticationType.OAuth && string.IsNullOrEmpty(password))
                {
                    // if we have no password we want it to use the oauth token cache instead
                    // this is signaled by saying we are multifactor
                    detail.UseMfa = true;
                }

                if (Guid.TryParse(clientId, out var azureAppId)) detail.AzureAdAppId = azureAppId;

                if (!string.IsNullOrEmpty(password))
                {
                    if (detail.NewAuthType == AuthenticationType.ClientSecret)
                    {
                        detail.SetClientSecret(credentials.UserName.Password);
                    }
                    else
                    {
                        detail.SetPassword(credentials.UserName.Password);
                    }
                }

                switch (DiscoverPurpose(o))
                {
                    case OrganizationPurpose.Production:
                        detail.EnvironmentColor = Color.Red;
                        detail.EnvironmentTextColor = Color.White;
                        detail.EnvironmentText = $"👎 - {o.FriendlyName}";
                        break;
                    case OrganizationPurpose.Test:
                        detail.EnvironmentColor = Color.Yellow;
                        detail.EnvironmentTextColor = Color.Black;
                        detail.EnvironmentText = $"✋ - {o.FriendlyName}";
                        break;
                    case OrganizationPurpose.Development:
                        detail.EnvironmentColor = Color.Green;
                        detail.EnvironmentTextColor = Color.White;
                        detail.EnvironmentText = $"👍 - {o.FriendlyName}";
                        break;
                }

                return detail;
            });
            var connectionFile = new CrmConnections(connectionFileName);
            connectionFile.Connections.AddRange(connections);
            connectionFile.SerializeToFile(connectionFilePath);

            Console.WriteLine($"Connection file generated successfully: \nPATH: {connectionFilePath}");
            Console.WriteLine("Press any key to exit: ");
            Console.ReadKey();
        }


        private static AuthenticationType GuessNewAuthType(ClientCredentials credential, string appId, OrganizationDetail o)
        {
            if (!string.IsNullOrEmpty(credential.Windows?.ClientCredential?.UserName))
            {
                return AuthenticationType.AD;
            }
            if (credential.UserName != null)
            {
                if (!string.IsNullOrEmpty(appId))
                {
                    if (Guid.TryParse(credential.UserName.UserName, out _))
                    {
                        return AuthenticationType.ClientSecret;
                    }
                    else
                    {
                        return AuthenticationType.OAuth;
                    }
                }
                else
                {
                    var url = new Uri(o.Endpoints[EndpointType.WebApplication]);
                    if (url.Host.EndsWith("dynamics.com"))
                    {
                        return AuthenticationType.Office365;
                    }
                    else
                    {
                        return AuthenticationType.IFD;
                    }
                }
            }
            return AuthenticationType.InvalidConnection;
        }

        private static Microsoft.Xrm.Sdk.Client.AuthenticationProviderType GuessAuthType(ClientCredentials credential, string appId, OrganizationDetail o)
        {
            var newAuthType = GuessNewAuthType(credential, appId, o);
            switch (newAuthType)
            {
                case AuthenticationType.AD:
                    return Microsoft.Xrm.Sdk.Client.AuthenticationProviderType.ActiveDirectory;
                case AuthenticationType.IFD:
                case AuthenticationType.Claims:
                    return Microsoft.Xrm.Sdk.Client.AuthenticationProviderType.Federation;
                case AuthenticationType.Live:
                case AuthenticationType.Office365:
                    return Microsoft.Xrm.Sdk.Client.AuthenticationProviderType.LiveId;
                case AuthenticationType.OAuth:
                case AuthenticationType.Certificate:
                case AuthenticationType.ClientSecret:
                case AuthenticationType.ExternalTokenManagement:
                case AuthenticationType.InvalidConnection:
                    return Microsoft.Xrm.Sdk.Client.AuthenticationProviderType.None;
            }
            return Microsoft.Xrm.Sdk.Client.AuthenticationProviderType.None;
        }

        public enum OrganizationPurpose
        {
            Production,
            Test,
            Development,
        }

        private static OrganizationPurpose DiscoverPurpose(OrganizationDetail organization)
        {
            foreach (var endpoint in organization.Endpoints.Values)
            {
                var uri = new Uri(endpoint);
                var lastSegment = uri.Segments.LastOrDefault();
                if (lastSegment != null)
                {
                    if (lastSegment.EndsWith("dev", StringComparison.CurrentCultureIgnoreCase)) return OrganizationPurpose.Development;
                    else if (lastSegment.EndsWith("develop", StringComparison.CurrentCultureIgnoreCase)) return OrganizationPurpose.Development;
                    else if (lastSegment.EndsWith("development", StringComparison.CurrentCultureIgnoreCase)) return OrganizationPurpose.Development;
                    else if (lastSegment.EndsWith("uat", StringComparison.CurrentCultureIgnoreCase)) return OrganizationPurpose.Test;
                    else if (lastSegment.EndsWith("test", StringComparison.CurrentCultureIgnoreCase)) return OrganizationPurpose.Test;
                }
                var deepestDomain = uri.Host.Split('.').FirstOrDefault();
                if (deepestDomain != null)
                {
                    if (deepestDomain.EndsWith("dev", StringComparison.CurrentCultureIgnoreCase)) return OrganizationPurpose.Development;
                    else if (deepestDomain.EndsWith("develop", StringComparison.CurrentCultureIgnoreCase)) return OrganizationPurpose.Development;
                    else if (deepestDomain.EndsWith("development", StringComparison.CurrentCultureIgnoreCase)) return OrganizationPurpose.Development;
                    else if (deepestDomain.EndsWith("uat", StringComparison.CurrentCultureIgnoreCase)) return OrganizationPurpose.Test;
                    else if (deepestDomain.EndsWith("test", StringComparison.CurrentCultureIgnoreCase)) return OrganizationPurpose.Test;
                }
            }
            return OrganizationPurpose.Production;
        }
    }
}
