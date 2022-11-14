# Exercise 2: Upload files to OneDrive

In this exercise, you'll learn how to upload files to OneDrive using Microsoft Graph, including both small and large files.

> [!IMPORTANT]
> This exercise assumes you have created the Azure AD application and .NET console application from the previous unit in this module. You'll edit the existing Azure AD application and .NET console application created in that exercise in this exercise.

## Task 1: Update the Azure AD application

The first step is to add a permission grant to the Azure AD application that will enable the .NET Core console app to upload files to the signed-in user's OneDrive account.

1. Open a browser and navigate to the [Azure Active Directory admin center (https://aad.portal.azure.com)](https://aad.portal.azure.com). Sign in using a **Work or School Account** that has global administrator rights to the tenancy.

2. Select **Azure Active Directory** in the left-hand navigation.

  ![Screenshot of the App registrations.](../../Linked_Image_Files/2-Graph/access-file-data/azure-ad-portal-home.png)

3. Select **Manage > App registrations** in the left-hand navigation.

4. On the **App registrations** page, select the **Graph Console App**.

5. Select **API Permissions** in the left-hand navigation panel.

6. Select the **Add a permission** button.

![Screenshot of the Add permission button.](../../Linked_Image_Files/2-Graph/access-file-data/05-azure-ad-portal-new-app-permissions-02.png)

7. In the **Request API permissions** panel that appears, select **Microsoft Graph** from the **Microsoft APIs** tab.

![Screenshot of Microsoft Graph in the Request API permissions panel.](../../Linked_Image_Files/2-Graph/access-file-data/azure-ad-portal-new-app-permissions-03.png)

8. When prompted for the type of permission, select **Delegated permissions**.

9. Enter **Files.ReadWrite** in the **Select permissions** search box and select the **Files.ReadWrite** permission, followed by the **Add permission** button at the bottom of the panel.

![Screenshot of the Files.ReadWrite permission in the Request API permissions panel.](../../Linked_Image_Files/2-Graph/access-file-data/05-azure-ad-portal-new-app-permissions-04.png)

10. In the **Configured Permissions** panel, select the button **Grant admin consent for [tenant]**, and then select the **Yes** button in the consent dialog to grant all users in your organization this permission.

## Task 2: Update .NET Core console application to upload a small file

In this section, you'll update the .NET console app to upload a small file to the user's OneDrive account.

1. Create a text file named **smallfile.txt** and fill it with a few paragraphs of text. Save this file to the root of the .NET Core console app's folder.

2. Locate the **Program.cs** file from the application you created in a previous unit in this module. Within the `Main` method, locate the following line:

```csharp
Console.WriteLine("Hello " + profileResponse.DisplayName);
```

3. Delete all code within the `Main` method after the above line.

4. Add the following code to the end of the `Main` method. This code will first get a reference to the file in the project. It will then open the file and upload it using the `PutAsync` method on the Microsoft Graph .NET SDK:

```csharp
// request 1 - upload small file to user's onedrive
var fileName = "smallfile.txt";
var filePath = Path.Combine(System.IO.Directory.GetCurrentDirectory(), fileName);
Console.WriteLine("Uploading file: " + fileName);

FileStream fileStream = new FileStream(filePath, FileMode.Open);
var uploadedFile = client.Me.Drive.Root
                              .ItemWithPath("smallfile.txt")
                              .Content
                              .Request()
                              .PutAsync<DriveItem>(fileStream)
                              .Result;
Console.WriteLine("File uploaded to: " + uploadedFile.WebUrl);
```

### Build and test the application

5. Run the following command in a command prompt to compile and run the console application:

```console
dotnet build
dotnet run
```

You now need to authenticate with Azure Active Directory. A new tab in your default browser should open to a page asking you to sign in. After you've logged in successfully, you'll be redirected to a page displaying the message, **"Authentication complete. You can return to the application. Feel free to close this browser tab"**. You may now close the browser tab and switch back to the console application.

6. The path to the uploaded file will be written to the console:

![Screenshot of the console application showing the uploaded file from the user's OneDrive.](../../Linked_Image_Files/2-Graph/access-file-data/05-app-run-01.png)

7. Open a browser and navigate to the URL written to the console, except omit the filename from the URL. After signing in using the same credentials used when testing the console app, you'll see the file listed in the user's OneDrive:

![Screenshot of the user's OneDrive account with 'smallfile.txt' highlighted.](../../Linked_Image_Files/2-Graph/access-file-data/05-app-run-02.png)

## Task 3: Update .NET Core console application to upload a large file

In this section, update the console app upload a large file.

1. Obtain a large file, one at least 5 MB, and copy it to the root of the .NET Core console app's folder. Name the file **largefile.zip**.

2. Locate the code you added above for `// request 1 - upload small file to user's onedrive` and comment it out so it doesn't continue to execute.

3. Add the following code to the `Main` method of the console application. This will get a reference to the large file you copied to the folder:

```csharp
// request 2 - upload large file to user's onedrive
var fileName = "largefile.zip";
var filePath = Path.Combine(System.IO.Directory.GetCurrentDirectory(), fileName);
Console.WriteLine("Uploading file: " + fileName);
```

4. Add the following code to the end of the `Main` method. This will open the file as a stream and create a new upload session using the Microsoft Graph .NET SDK:

```csharp
// load resource as a stream
using (Stream stream = new FileStream(filePath, FileMode.Open))
{
  var uploadSession = client.Me.Drive.Root
                                  .ItemWithPath(fileName)
                                  .CreateUploadSession()
                                  .Request()
                                  .PostAsync()
                                  .Result;

  // create upload task

  // create progress implementation

  // upload file
}
```

5. Next, add the following code after the comment `// create upload task`. This code will create a new `LargeFileUploadTask` using the existing session and stream:

```csharp
var maxChunkSize = 320 * 1024;
var largeUploadTask = new LargeFileUploadTask<DriveItem>(uploadSession, stream, maxChunkSize);
```

> [!NOTE]
> If your app splits a file into multiple byte ranges, the size of each byte range MUST be a multiple of 320 KiB (327,680 bytes). Using a fragment size that does not divide evenly by 320 KiB will result in errors committing some files.

6. The `LargeFileUploadTask` object will report the progress of the file upload. To use this capability, create an implementation of the `IProgress` interface. Add the following code after the `// create progress implementation` comment:

```csharp
IProgress<long> uploadProgress = new Progress<long>(uploadBytes =>
{
  Console.WriteLine($"Uploaded {uploadBytes} bytes of {stream.Length} bytes");
});
```

7. Finally, upload the file by adding the following code after the `// upload file` comment:

```csharp
UploadResult<DriveItem> uploadResult = largeUploadTask.UploadAsync(uploadProgress).Result;
if (uploadResult.UploadSucceeded)
{
  Console.WriteLine("File uploaded to user's OneDrive root folder.");
}
```

### Build and test the application

8. Run the following command in a command prompt to compile and run the console application:

```console
dotnet build
dotnet run
```

9. After you've signed in, the console app will write the progress of uploading the file until it's complete:

![Screenshot of the console application showing uploading files to the user's OneDrive.](../../Linked_Image_Files/2-Graph/access-file-data/05-app-run-03.png)

10. Open a browser and navigate to the user's OneDrive account to see the large file that has been uploaded:

![Screenshot of the user's OneDrive account with 'largefile.zip' highlighted.](../../Linked_Image_Files/2-Graph/access-file-data/05-app-run-04.png)

## Summary

In this exercise, you learned how to upload files to OneDrive using Microsoft Graph, including both small and large files.
