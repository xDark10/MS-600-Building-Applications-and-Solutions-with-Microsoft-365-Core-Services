# Exercise 1: Understand throttling in Microsoft Graph

In this exercise, you'll create a new Azure AD web application registration using the Azure Active Directory admin center, a .NET Core console application and query Microsoft Graph. You'll issue many requests in parallel to trigger your requests to be throttled. This application will allow you to see the response you'll receive.

## Prerequisites

Developing Microsoft Graph apps requires a Microsoft 365 tenant.

For the Microsoft 365 tenant, follow the instructions on the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program) site for obtaining a developer tenant if you don't currently have a Microsoft 365 account.

You'll use the .NET SDK to create custom Microsoft Graph app in this module. The exercises in this module assume you have the following tools installed on your developer workstation.

> [!IMPORTANT]
> In most cases, installing the latest version of the following tools is the best option. The versions listed here were used when this module was published and last tested.

- [.NET SDK](https://dotnet.microsoft.com/) - v5.\* (or higher)
- [Visual Studio Code](https://code.visualstudio.com)

You must have the minimum versions of these prerequisites installed on your workstation.

## Task 1: Create an Azure AD application

1. Open a browser and navigate to the [Azure Active Directory admin center (https://aad.portal.azure.com)](https://aad.portal.azure.com). Sign in using a **Work or School Account** that has global administrator rights to the tenancy.

2. Select **Azure Active Directory** in the left-hand navigation.

  ![Screenshot of the App registrations](../../Linked_Image_Files/02-02-azure-ad-portal-home.png)

3. Select **Manage > App registrations** in the left-hand navigation.

4. On the **App registrations** page, select **+ New registration**.

  ![Screenshot of App Registrations page](../../Linked_Image_Files/02-02-azure-ad-portal-new-app-00.png)

5. On the **Register an application** page, set the values as follows:

- **Name**: Graph Console App
- **Supported account types**: Accounts in this organizational directory only (Contoso only - Single tenant)

    ![Screenshot of the Register an application page](../../Linked_Image_Files/02-02-azure-ad-portal-new-app-01.png)

6. Select **Register**.

7. On the **Graph Console App** page, copy the value of the **Application (client) ID** and **Directory (tenant) ID**; you'll need these in the application.

  ![Screenshot of the application ID of the new app registration](../../Linked_Image_Files/02-02-azure-ad-portal-new-app-details.png)

8. Select **Manage > Authentication**.

9. In the **Platform configurations** section, select the **+ Add a platform** button. Then in the **Configure platforms** panel, select the **Mobile and desktop applications** button:

![Screenshot of the Platform configurations section](../../Linked_Image_Files/02-02-azure-ad-portal-new-app-02.png)

10. In the **Redirect URIs** section of the **Configure Desktop + devices** panel, select the entry that ends with **nativeclient**, and then select the **Configure** button:

![Screenshot of the Configure Desktop + devices panel](../../Linked_Image_Files/02-02-azure-ad-portal-new-app-03.png)

11. In the **Authentication** page, scroll down to the **Allow public client flows** section and set the toggle to **Yes**.

![Screenshot of the Default client type section](../../Linked_Image_Files/02-02-azure-ad-portal-new-app-04.png)

12. Select **Save** in the top menu to save your changes.

### Grant Azure AD application permissions to Microsoft Graph

After creating the application, you need to grant it the necessary permissions to Microsoft Graph

13. Select **Manage > API Permissions** in the left-hand navigation panel.

![Screenshot of the API Permissions navigation item](../../Linked_Image_Files/02-02-azure-ad-portal-new-app-permissions-01.png)

14. Select the **+ Add a permission** button.

15. In the **Request API permissions** panel that appears, select **Microsoft Graph** from the **Microsoft APIs** tab.

![Screenshot of Microsoft Graph in the Request API permissions panel](../../Linked_Image_Files/02-02-azure-ad-portal-new-app-permissions-02.png)

16. When prompted for the type of permission, select **Delegated permissions**.

![Screenshot of the Mail.Read permission in the Request API permissions panel](../../Linked_Image_Files/02-02-azure-ad-portal-new-app-permissions-03.png)

17. Enter **Mail.R** in the **Select permissions** search box and select the **Mail.Read** permission, followed by the **Add permissions** button at the bottom of the panel.

18. In the **Configured Permissions** panel, select the button **Grant admin consent for [tenant]**, and then select **Yes** in the confirmation dialog.

![Screenshot of the Configured permissions panel](../../Linked_Image_Files/02-02-azure-ad-portal-new-app-permissions-04.png)

> [!NOTE]
> The option to **Grant admin consent** here in the Azure AD admin center is pre-consenting the permissions to the users in the tenant to simplify the exercise. This approach allows the console application to use the [resource owner password credential grant](/azure/active-directory/develop/v2-oauth-ropc), so the user isn't prompted to grant consent to the application that simplifies the process of obtaining an OAuth access token. You could elect to implement alternative options such as the [device code flow](/azure/active-directory/develop/v2-oauth2-device-code) to utilize dynamic consent as another option.

## Task 2: Create .NET Core console application


1. Open your command prompt, navigate to a directory where you have rights to create your project, and run the following command to create a new .NET Core console application:

```console
dotnet new console -o graphconsoleapp
```

2. After creating the application, run the following commands to ensure your new project runs correctly.

```console
cd graphconsoleapp
dotnet add package Microsoft.Identity.Client
dotnet add package Microsoft.Graph
dotnet add package Microsoft.Extensions.Configuration
dotnet add package Microsoft.Extensions.Configuration.FileExtensions
dotnet add package Microsoft.Extensions.Configuration.Json
```

3. Open the application in Visual Studio Code using the following command:

```console
code .
```

4. If Visual Studio code displays a dialog box asking if you want to add required assets to the project, select **Yes**.

### Update the console app to enable nullable reference types

Nullable reference types refers to a group of features introduced in C# 8.0 that you can use to minimize the likelihood that your code causes the runtime to throw System.NullReferenceException.

Nullable reference types are enabled by default in .NET 6 projects, they are disabled by default in .NET 5 projects.

Ensuring that nullable reference types are enabled is not related to the use of Microsoft Graph, it just ensures the exercises in this module can contain a single set of code that will compile without warnings when using either .NET 5 or .NET 6.

Open the **graphconsoleapp.csproj** file and ensure the `<PropertyGroup>` element contains the following child element:

```xml
<Nullable>enable</Nullable>
```

### Update the console app to support Azure AD authentication

5. Create a new file named **appsettings.json** in the root of the project and add the following code to it:

```json
{
  "tenantId": "YOUR_TENANT_ID_HERE",
  "applicationId": "YOUR_APP_ID_HERE"
}
```

6. Update properties with the following values:

- `YOUR_TENANT_ID_HERE`: Azure AD directory ID
- `YOUR_APP_ID_HERE`: Azure AD client ID

#### Create helper classes

7. Create a new folder **Helpers** in the project.

8. Create a new file **AuthHandler.cs** in the **Helpers** folder and add the following code:

```csharp
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace Helpers
{
  public class AuthHandler : DelegatingHandler
  {
    private IAuthenticationProvider _authenticationProvider;

    public AuthHandler(IAuthenticationProvider authenticationProvider, HttpMessageHandler innerHandler)
    {
      InnerHandler = innerHandler;
      _authenticationProvider = authenticationProvider;
    }

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
      await _authenticationProvider.AuthenticateRequestAsync(request);
      return await base.SendAsync(request, cancellationToken);
    }
  }
}
```

9. Create a new file **MsalAuthenticationProvider.cs** in the **Helpers** folder and add the following code:

```csharp
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;

namespace Helpers
{
  public class MsalAuthenticationProvider : IAuthenticationProvider
  {
    private static MsalAuthenticationProvider? _singleton;
    private IPublicClientApplication _clientApplication;
    private string[] _scopes;
    private string _username;
    private SecureString _password;
    private string? _userId;

    private MsalAuthenticationProvider(IPublicClientApplication clientApplication, string[] scopes, string username, SecureString password)
    {
      _clientApplication = clientApplication;
      _scopes = scopes;
      _username = username;
      _password = password;
      _userId = null;
    }

    public static MsalAuthenticationProvider GetInstance(IPublicClientApplication clientApplication, string[] scopes, string username, SecureString password)
    {
      if (_singleton == null)
      {
        _singleton = new MsalAuthenticationProvider(clientApplication, scopes, username, password);
      }

      return _singleton;
    }

    public async Task AuthenticateRequestAsync(HttpRequestMessage request)
    {
      var accessToken = await GetTokenAsync();

      request.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
    }

    public async Task<string> GetTokenAsync()
    {
      if (!string.IsNullOrEmpty(_userId))
      {
        try
        {
          var account = await _clientApplication.GetAccountAsync(_userId);

          if (account != null)
          {
            var silentResult = await _clientApplication.AcquireTokenSilent(_scopes, account).ExecuteAsync();
            return silentResult.AccessToken;
          }
        }
        catch (MsalUiRequiredException) { }
      }

      var result = await _clientApplication.AcquireTokenByUsernamePassword(_scopes, _username, _password).ExecuteAsync();
      _userId = result.Account.HomeAccountId.Identifier;
      return result.AccessToken;
    }
  }
}
```

### Incorporate Microsoft Graph into the console app

10. Open the **Program.cs** file and add the following `using` statements to the top fo the file:
    
```csharp
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Helpers;

namespace graphconsoleapp
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
        }
    }
}
```

11. Add the following method `LoadAppSettings` to the `Program` class. The method retrieves the configuration details from the **appsettings.json** file previously created:

```csharp
private static IConfigurationRoot? LoadAppSettings()
{
  try
  {
    var config = new ConfigurationBuilder()
                      .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                      .AddJsonFile("appsettings.json", false, true)
                      .Build();

    if (string.IsNullOrEmpty(config["applicationId"]) ||
        string.IsNullOrEmpty(config["tenantId"]))
    {
      return null;
    }

    return config;
  }
  catch (System.IO.FileNotFoundException)
  {
    return null;
  }
}
```

12. Add the following method `CreateAuthorizationProvider` to the `Program` class. The method will create an instance of the clients used to call Microsoft Graph.

```csharp
private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config, string userName, SecureString userPassword)
{
  var clientId = config["applicationId"];
  var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

  List<string> scopes = new List<string>();
  scopes.Add("User.Read");
  scopes.Add("Mail.Read");

  var cca = PublicClientApplicationBuilder.Create(clientId)
                                          .WithAuthority(authority)
                                          .Build();
  return MsalAuthenticationProvider.GetInstance(cca, scopes.ToArray(), userName, userPassword);
}
```

13. Add the following method `GetAuthenticatedHTTPClient` to the `Program` class. The method creates an instance of the `HttpClient` object.

```csharp
private static HttpClient GetAuthenticatedHTTPClient(IConfigurationRoot config, string userName, SecureString userPassword)
{
  var authenticationProvider = CreateAuthorizationProvider(config, userName, userPassword);
  var httpClient = new HttpClient(new AuthHandler(authenticationProvider, new HttpClientHandler()));
  return httpClient;
}
```

14. Add the following method `ReadPassword` to the `Program` class. The method prompts the user for their password:

```csharp
private static SecureString ReadPassword()
{
  Console.WriteLine("Enter your password");
  SecureString password = new SecureString();
  while (true)
  {
    ConsoleKeyInfo c = Console.ReadKey(true);
    if (c.Key == ConsoleKey.Enter)
    {
      break;
    }
    password.AppendChar(c.KeyChar);
    Console.Write("*");
  }
  Console.WriteLine();
  return password;
}
```

15. Add the following method `ReadUsername` to the `Program` class. The method prompts the user for their username:

```csharp
private static string ReadUsername()
{
  string? username;
  Console.WriteLine("Enter your username");
  username = Console.ReadLine();
  return username ?? "";
}
```

16. Locate the `Main` method in the `Program` class. Add the following code to the end of the `Main` method to load the configuration settings from the **appsettings.json** file:

```csharp
var config = LoadAppSettings();
if (config == null)
{
  Console.WriteLine("Invalid appsettings.json file.");
  return;
}
```

17. Add the following code to the end of the `Main` method, just after the code added in the last step. This code will obtain an authenticated instance of the `HttpClient` and submit a request for the current user's email:

```csharp
var userName = ReadUsername();
var userPassword = ReadPassword();

var client = GetAuthenticatedHTTPClient(config, userName, userPassword);
```

18. Add the following code to issue many requests to Microsoft Graph. This code will create a collection of tasks to request a specific Microsoft Graph endpoint. When a task succeeds, it will write a dot to the console while failed request will write an `X` to the console. The most recent failed request's status code and headers are saved.

19. All tasks are then executed in parallel. At the conclusion of all requests, the results are written to the console:

```csharp
var totalRequests = 100;
var successRequests = 0;
var tasks = new List<Task>();
var failResponseCode = HttpStatusCode.OK;
HttpResponseHeaders failedHeaders = null!;

for (int i = 0; i < totalRequests; i++)
{
  tasks.Add(Task.Run(() =>
  {
    var response = client.GetAsync("https://graph.microsoft.com/v1.0/me/messages").Result;
    Console.Write(".");
    if (response.StatusCode == HttpStatusCode.OK)
    {
      successRequests++;
    }
    else
    {
      Console.Write('X');
      failResponseCode = response.StatusCode;
      failedHeaders = response.Headers;
    }
  }));
}

var allWork = Task.WhenAll(tasks);
try
{
  allWork.Wait();
}
catch { }
Console.WriteLine();
Console.WriteLine("{0}/{1} requests succeeded.", successRequests, totalRequests);
if (successRequests != totalRequests)
{
  Console.WriteLine("Failed response code: {0}", failResponseCode.ToString());
  Console.WriteLine("Failed response headers: {0}", failedHeaders);
}
```

### Build and test the application

20. Run the following command to ensure the developer certificate has been trusted:

```console
dotnet dev-certs https --trust
```

21. Run the following command in a command prompt to compile the console application:

```console
dotnet build
```

22. Run the following command to run the console application:

```console
dotnet run
```

> [!TIP]
> The console app may take a one or two minutes to complete the process of authenticating and obtaining an access token from Azure AD and issuing the requests to Microsoft Graph.

23. After entering the username and password of a user, you'll see the results written to the console:

![Screenshot of the console application with no query parameters](../../Linked_Image_Files/02-02-03-app-run-01.png)

There's a mix of success and failure indicators in the console. The summary states only 39% of the requests were successful.

After the results, the console has two lines that begin with **Failed response**. Notice the code states **TooManyRequests** that is the representation of the HTTP status code 429. This status code is the indication that your requests are being throttled.

Also notice within the collection of response headers, the presence of **Retry-After**. This header is the value in seconds that Microsoft Graph tells you to wait before sending your next request to avoid being further throttled.

## Summary

In this exercise, you created an Azure AD application and .NET console application that retrieved user data from Microsoft Graph. This query was submitted multiple times to demonstrate Microsoft Graph throttling the requests.
