# Exercise 3: Eliminate polling Microsoft Graph with delta query

In this exercise, you'll use the Azure AD application and .NET console application you previously created and modify them to demonstrate using the delta query feature of Microsoft Graph. An application can use the delta query feature to avoid costly requests to poll for changes in Microsoft Graph that can trigger requests to be throttled.

> [!IMPORTANT]
> This exercise assumes you have created the Azure AD application and .NET console application from the previous unit in this module. You'll edit the existing Azure AD application and .NET console application created in that exercise in this exercise.

## Task 1: Add more permissions the Azure AD application

In this exercise, you'll update the application to get a list of all the users in the tenant. To do this task, the Azure AD application needs more permissions.

1. Open a browser and navigate to the [Azure Active Directory admin center (https://aad.portal.azure.com)](https://aad.portal.azure.com). Sign in using a **Work or School Account**.

2. Select **Azure Active Directory** in the left-hand navigation.

  ![Screenshot of Azure AD Portal](../../Linked_Image_Files/2-Graph/optimize-network-traffic/azure-ad-portal-home.png)

3. Select **Manage > App registrations** in the left-hand navigation.

4. On the **App registrations** page, select the **Graph Console App**, or the name of the application you created in this module in a previous exercise.

5. Select **API Permissions** in the left-hand navigation panel.

![Screenshot of the API Permissions navigation item](../../Linked_Image_Files/2-Graph/optimize-network-traffic/07-azure-ad-portal-add-permission.png)

6. Select the **Add a permission** button.

7. In the **Request API permissions** panel that appears, select **Microsoft Graph** from the **Microsoft APIs** tab.

![Screenshot of Microsoft Graph in the Request API permissions panel](../../Linked_Image_Files/2-Graph/optimize-network-traffic/azure-ad-portal-new-app-permissions-02.png)

8. When prompted for the type of permission, select **Delegated permissions** and search for the permission **User.Read.All**, select it and then select the **Add permission** button at the bottom of the panel.

9. In the **Configured Permissions** panel, select the button **Grant admin consent for [tenant]**, and then select the **Yes** button in the consent dialog to grant all users in your organization this permission.

## Task 2: Update the console application to use delta query

Now you'll update the application to retrieve and then display all users in the tenant with Microsoft Graph. After displaying the list of users, the application will sleep for 10 seconds and repeat the same query, but instead only display the users who have been added or updated.

1. First, add the following two members to the class, immediately before the `Main` method:

```csharp
private static object? _deltaLink = null;
private static IUserDeltaCollectionPage? _previousPage = null;
```

### Add a Method to Display Users

2. Within the **Program.cs** file, add the following method to write all users to the console:

```csharp
private static void OutputUsers(IUserDeltaCollectionPage users)
{
  foreach (var user in users)
  {
    Console.WriteLine($"User: {user.Id}, {user.GivenName} {user.Surname}");
  }
}
```

## Task 3: Add a method to retrieve users

Next, add a method to retrieve users with Microsoft Graph. This method will use the delta link feature in Microsoft Graph.

If this is the first time the request is made, indicated by the `_previousPage == null`, you'll request all users but also request a delta link using the `Delta()` method. The code will return the current page of users.

Otherwise, if there are more pages of users, it will start a request for the next page using the delta link string that you'll obtain later.

1. Add the following method to the `Program` class:

```csharp
private static IUserDeltaCollectionPage GetUsers(GraphServiceClient graphClient, object? deltaLink)
{
  IUserDeltaCollectionPage page;

  // IF this is the first request, then request all users
  //    and include Delta() to request a delta link to be included in the
  //    last page of data
  if (_previousPage == null || deltaLink == null)
  {
      page = graphClient.Users
                        .Delta()
                        .Request()
                        .Select("Id,GivenName,Surname")
                        .GetAsync()
                        .Result;
  }
  // ELSE, not the first page so get the next page of users
  else
  {
    _previousPage.InitializeNextPageRequest(graphClient, deltaLink.ToString());
    page = _previousPage.NextPageRequest.GetAsync().Result;
  }

  _previousPage = page;
  return page;
}
```

## Task 4: Add a method to check for new and changed users

The next method will be used to check for new and changed users. The first time it runs, it will get all users and display them page by page until it reaches the last page of the response from Microsoft Graph.

1. Add the following method to the `Program` class:

```csharp
private static void CheckForUpdates(IConfigurationRoot config)
{
  var graphClient = GetAuthenticatedGraphClient(config);

  // get a page of users
  var users = GetUsers(graphClient, _deltaLink);

  OutputUsers(users);

  // go through all of the pages so that we can get the delta link on the last page.
  while (users.NextPageRequest != null)
  {
    users = users.NextPageRequest.GetAsync().Result;
    OutputUsers(users);
  }
}
```

The last thing this method needs to do is retrieve the delta link from the last page of results. This delta link, a string, will be used in future requests that will signal to Microsoft Graph to only return users who have been added or changed since the previous request.

2. Add the following code to the end of the `CheckForUpdates` method:

```csharp
object? deltaLink;

if (users.AdditionalData.TryGetValue("@odata.deltaLink", out deltaLink))
{
  _deltaLink = deltaLink;
}
```

## Task 5: Update the application to use the new `CheckForUpdates` method

1. Locate the following code in the `Main` method:

```csharp
var config = LoadAppSettings();
if (config == null)
{
  Console.WriteLine("Invalid appsettings.json file.");
  return;
}
```

2. Remove all the existing code in this method after these lines.

3. Next, add the following code to after the previous two lines above. This will use the `CheckForUpdates` method to get a list of all users. It will then go into an infinite loop, sleeping for 10 seconds, and issue the same request again:

```csharp
Console.WriteLine("All users in tenant:");
CheckForUpdates(config);
Console.WriteLine();
while (true)
{
  Console.WriteLine("... sleeping for 10s - press CTRL+C to terminate");
  System.Threading.Thread.Sleep(10 * 1000);
  Console.WriteLine("> Checking for new/updated users since last query...");
  CheckForUpdates(config);
}
```

The first request will result in the application obtaining the delta link that will be used in future requests to include only the new and changed users. Each time users are requested, the delta link will be refreshed.

### Build and test the application

4. Run the following command in a command prompt to compile the console application:

```console
dotnet build
```

5. Run the following command to run the console application:

```console
dotnet run
```

You now need to authenticate with Azure Active Directory. A new tab in your default browser should open to a page asking you to sign-in. After you've logged in successfully, you'll be redirected to a page displaying the message, **"Authentication complete. You can return to the application. Feel free to close this browser tab"**. You may now close the browser tab and switch back to the console application.

6. The application will write all the users in the tenant to the console:

![Screenshot of the console displaying all users](../../Linked_Image_Files/2-Graph/optimize-network-traffic/07-app-run-01.png)

7. If you let the app run for a few moments without doing anything, notice it will keep requesting users, but not display anyone as no users have been created in your tenant:

![Screenshot of the console displaying no new users](../../Linked_Image_Files/2-Graph/optimize-network-traffic/07-app-run-02.png)

8. Now, add a new user to your tenant.

9. Open a browser and navigate to the [Azure Active Directory admin center (https://aad.portal.azure.com)](https://aad.portal.azure.com). Sign in using a **Work or School Account**.

10. Select **Users** in the left-hand navigation.

![Screenshot of the Azure AD Users screen](../../Linked_Image_Files/2-Graph/optimize-network-traffic/azure-ad-portal-users.png)

11. On the **Users - All Users** page, select **New User**.

12. Create a new user by entering the values in the **Identity** section of the **New User** page and select **Create** at the bottom of the form.

![Screenshot creating a new user](../../Linked_Image_Files/2-Graph/optimize-network-traffic/07-app-run-03.png)

13. Watch the running application in the console. The next time it runs, it will display the new user that you created:

![Screenshot of the console app displaying the new user](../../Linked_Image_Files/2-Graph/optimize-network-traffic/07-app-run-04.png)

14. Press <kbd>CTRL</kbd>+<kbd>C</kbd> to stop the application.

## Summary

In this exercise, you used the Azure AD application and .NET console application you previously created and modified them to demonstrate using the delta query feature of Microsoft Graph. An application can use the delta query feature to avoid costly requests to poll for changes in Microsoft Graph that can trigger requests to be throttled.
