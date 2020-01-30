# Exercise 2: Understanding webhooks in Microsoft Teams

You use webhooks to connect web services to channels and teams in Teams.

You can create two types of webhooks—incoming and outgoing. Incoming webhooks allow you to connect a channel to a service, and outgoing webhooks allow you to send messages to a service.

## Task 1: Setting up a custom incoming webhook

These steps explain how to send a card to a connector.

1. In Microsoft Teams, select the **Apps** icon located in the left navigation. Select **Connectors** then select **Incoming Webhook**.

    ![Apps to Connectors to Incoming webhook](../../Linked_Image_Files/m04_e02_t01_image_1.png)

1. From the dialog, select **Add to a team**.

    ![Incoming webhook dialog with Add to a team button.](../../Linked_Image_Files/m04_e02_t01_image_2.png)

1. From the **Select a channel to start using Incoming Webhook**, select the **General** channel of the **Training** **Content** team.

    ![Select a channel to start using incoming webhook.](../../Linked_Image_Files/m04_e02_t01_image_3.png)

1. Select the **Set up a connector** button, provide a name, and, optionally, upload an image avatar for your webhook then select **Create**.

    ![Incoming webhook configuration dialog.](../../Linked_Image_Files/m04_e02_t01_image_4.png)

1. The dialog window will present a unique URL that will map to the channel. Make sure that you **copy and save the URL**—you will need to provide it to the outside service.

    ![Incoming webhook configuration dialog.](../../Linked_Image_Files/m04_e02_t01_image_5.png)

1. Select the **Done** button. The webhook will be available in the team channel.

1. The Incoming Webhook should now be available in the **Configured** section.

![Incoming webhook screen](../../Linked_Image_Files/m04_e02_t01_image_6.png)

## Task 2: Post a message to the webhook using PowerShell

1. From the **PowerShell** prompt, enter the following command and replace the `<YOUR WEBHOOK URL>` with the URL you copied from the previous steps when you setup your incoming webhook.

    ```powershell
    Invoke-RestMethod -Method post -ContentType 'Application/Json' -Body '{"text":"Hello World!"}' -Uri <YOUR WEBHOOK URL>
    ```

1. If the POST succeeds, you should see a simple 1 output by `Invoke-RestMethod`.

    ![Invoke rest method PowerShell command executing.](../../Linked_Image_Files/m04_e02_t02_image_1.png)

1. Check the Microsoft Teams channel associated with the webhook URL. You should see the new card posted to the channel.

![My Incoming Webhook hello world message in Teams channel.](../../Linked_Image_Files/m04_e02_t02_image_2.png)

## Review

In this exercise, you:

- Used a combination of coding tools and user interfaces to create incoming webhook.

