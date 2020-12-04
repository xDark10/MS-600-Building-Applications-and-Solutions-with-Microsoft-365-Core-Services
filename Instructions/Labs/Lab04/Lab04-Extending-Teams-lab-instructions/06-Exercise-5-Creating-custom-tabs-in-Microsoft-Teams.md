# Exercise 5: Creating custom tabs in Microsoft Teams

## Task 1: Create a personal tab

This task explains how to use Visual Studio Code and App Studio to create a personal tab. For this task, you are going to modify the app you created in the first exercise.

### Create your personal tab app project

1. In **Visual Studio Code**, select **Microsoft Teams** on the left Activity Bar and choose **Create a new Teams app**.

    ![Create a new Teams app](../../Linked_Image_Files/m04_e01_t02_image_1.png)

1. When prompted, sign in with your Microsoft 365 development account.

1. On the **Add capabilities** screen, select **Tab** then Next.

    ![Add capabilities](../../Linked_Image_Files/m04_e01_t02_image_2.png)

1. Enter a name for your Teams app. (This is the default name for your app and also the name of the app project directory on your local machine.)

1. Check only the **Personal tab** option and select **Finish** at the bottom of the screen to configure your project.

### Build and run the personal tab app

1. Open Terminal in Visual Studio Code. From the Visual Studio Code ribbon select **Terminal > New Terminal**.

1. Go to the root directory of your app project and run `npm install`.

1. Run `npm start`.

### Sideload the personal tab app in Teams

1. In Visual Studio Code, press the **F5** key to launch a Teams web client.

1. To display your app content in Teams, specify that where your app is running (localhost) is trustworthy:

    1. Open a new tab in the same browser window (Google Chrome by default) which opened after pressing F5.

    1. Go to https://localhost:3000/tab and proceed to the page.

1. Go back to Teams. In the dialog, select Add for me to install your app.

1. Please keep running this app in order to complete the next task.

## Task 2: Customize your personal tab

### Update the personal tab content page

1. Go to the **src/components** directory and open **Tab.js**. Locate the **render()** function and paste your content inside return() (as shown).

    ```html
    <div>
        <h1>Important Contacts</h1>
        <ul>
            <li>Help Desk: <a href="mailto:support@company.com">support@company.com</a></li>
            <li>Human Resources: <a href="mailto:hr@company.com">hr@company.com</a></li>
            <li>Facilities: <a href="mailto:facilities@company.com">facilities@company.com</a></li>
        </ul>
    </div>
    ```

1. Add the following rule to **App.css** so the email links are easier to read no matter which theme is used.

    ```css
    a {
        color: inherit;
    }
    ```

1. Save your changes. Go to your app's tab in Teams to view the new content.

    ![Add capabilities](../../Linked_Image_Files/personal-tab-tutorial-content.png)

### Update the personal tab theme

Good apps feel native to Teams, so it's important your tab blends with the Teams theme your users prefer: default (light), dark, or high contrast. As you might have noticed in the last screenshot, your tab still has a light background when the client's using the dark theme. This is not a recommended user experience.

1. Open tab.js with Visual Studio Code, and then insert the following theme change handler immediately after the `microsoftTeams.getContext()` call.

  ```typescript
  microsoftTeams.registerOnThemeChangeHandler(theme => {
    if (theme !== this.state.theme) {
      this.setState({ theme });  
    }
  });
  ```

1. Replace `render()` method with the following code.

  ```typescript
   render() {
      const isTheme = this.state.theme;

      let newTheme;

      if (isTheme === "default") {
        newTheme = {
          backgroundColor: "#EEF1F5",
          color: "#16233A"
        };
      } else {
        newTheme = {
          backgroundColor: "#2B2B30",
          color: "#FFFFFF"
        };
      }

      return (
        <div style={newTheme}>
          <h1>Important Contacts</h1>
          <ul>
              <li>Help Desk: <a href="mailto:support@company.com">support@company.com</a></li>
              <li>Human Resources: <a href="mailto:hr@company.com">hr@company.com</a></li>
              <li>Facilities: <a href="mailto:facilities@company.com">facilities@company.com</a></li>
          </ul>
      </div>
      );
  }
   ```

1. Check your tab in Teams. The appearance should closely match the dark theme.

Well done! You have a Teams app with a personal tab that makes it easier to find important contacts in your organization.

## Task 3: Create a group/Teams channel tab

In this task, you'll build a basic channel tab (also known as a group tab), which is a full-screen page for a team channel or chat. Unlike a personal tab, users can configure some aspects of this kind of tab (for example, rename the tab so it's meaningful to their channel).

### Create your group tab app project

1. In **Visual Studio Code**, select **Microsoft Teams** on the left Activity Bar and choose **Create a new Teams app**.

    ![Create a new Teams app](../../Linked_Image_Files/m04_e01_t02_image_1.png)

1. When prompted, sign in with your Microsoft 365 development account.

1. On the **Add capabilities** screen, select **Tab** then Next.

    ![Add capabilities](../../Linked_Image_Files/m04_e01_t02_image_2.png)

1. Enter a name for your Teams app. (This is the default name for your app and also the name of the app project directory on your local machine.)

1. Check the **Personal tab** and **Group or Teams channel tab** options.

1. Select **Finish** at the bottom of the screen to configure your project.

## Task 4: Customize your group tab

### Update the group tab content page

1. Go to the **src/components** directory and open **Tab.js**. Locate the **render()** function and paste your content inside return() (as shown).

    ```html
    <div>
        <h1>Important Contacts</h1>
        <ul>
            <li>Help Desk: <a href="mailto:support@company.com">support@company.com</a></li>
            <li>Human Resources: <a href="mailto:hr@company.com">hr@company.com</a></li>
            <li>Facilities: <a href="mailto:facilities@company.com">facilities@company.com</a></li>
        </ul>
    </div>
    ```

1. Add the following rule to **App.css** so the email links are easier to read no matter which theme is used.

    ```css
    a {
        color: inherit;
    }
    ```

### Update the tab configuration page

1. Go to the **src/components** directory and open **TabConfig.js**. Locate the **render()** function and paste your content inside return() (as shown).

    ```javascript
    <div>
        <h1>Add My Contoso Contacts</h1>
        <div>
            Select <b>Save</b> to add our organization''s important contacts to this workspace.
        </div>
    </div>
    ```

1. In **TabConfig.js**, go to `microsoftTeams.settings.setSettings`. Add the suggestedDisplayName property with the tab name you want to display by default (as shown). Use the provided name or create your own. (By default, users to change the name if they want.)

    ```javascript
    microsoftTeams.settings.setSettings({
        "contentUrl": "https://localhost:3000/tab",
        "suggestedDisplayName": "Team Contacts"
    });
    ```

### Build and run the group tab app

1. Open Terminal in Visual Studio Code. From the Visual Studio Code ribbon select **Terminal > New Terminal**.

1. Go to the root directory of your app project and run `npm install`.

1. Run `npm start`.

### Sideload the group tab in Teams

1. In Visual Studio Code, press the **F5** key to launch a Teams web client.

1. To display your app content in Teams, specify that where your app is running (localhost) is trustworthy:

    1. Open a new tab in the same browser window (Google Chrome by default) which opened after pressing F5.

    1. Go to https://localhost:3000/tab and proceed to the page.

1. Go back to Teams. In the dialog, select **Add to a team** or **Add to a chat** and locate a channel or chat you can use for testing.

1. Select **Set up a tab**. The configuration page displays in a dialog.

    ![Screenshot of a channel tab configuration page.](../../Linked_Image_Files/channel-tab-tutorial-content.png)

1. Select Save to configure the tab. The content page displays.

    ![SScreenshot of a channel tab with static content view.](../../Linked_Image_Files/channel-tab-tutorial-content-installed.png)

Well done! You have a Teams app with a tab for displaying useful content in channels and chats.

## Review

In this exercise, you:

- Created a personal tab.

  - Created a personal tab app project.

  - Customize the personal tab.

  - Installed and tested the tab.

- Created a group/Teams channel tab.

  - Created a group/Teams channel tab app project.

  - Customize the group/Teams channel tab.

  - Customize tab configuration page.
  
  - Installed and tested the tab.
  