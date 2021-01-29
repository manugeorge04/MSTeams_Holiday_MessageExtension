# Hoiday Messaging Extension

A simple Messaging Extension developed for MS Teams that can retrieve and display the list of holdiays for any region from a database stored within Azure Cosmos DB. The messaging extension opens a Task Module on launch from which the user can choose a desired location from the given searchable dropdown list. The source code for the web app displayed within the Task Module can be found [here](https://github.com/manugeorge04/SinglePageMERNWebApp). Subsequently the user can opt to share the table within the chat with another end user as a card.

## Prerequisites

**Dependencies**
-  [NodeJS](https://nodejs.org/en/)
-  [ngrok](https://ngrok.com/) or equivalent tunneling solution
-  [M365 developer account](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/prepare-your-o365-tenant) or access to a Teams account with the appropriate permissions to install an app.

**Configure Ngrok**

Your app will be run from a localhost server. You will need to setup Ngrok in order to tunnel from the Teams client to localhost. 

**Run Ngrok**

Run ngrok - point to port 3978

`ngrok http -host-header=rewrite 3978`

**Update Bot Framework Messaging Endpoint**

  Note: You can also do this with the Manifest Editor in App Studio if you are familiar with the process.

- For the Messaging endpoint URL, use the current `https` URL you were given by running ngrok and append it with the path `/api/messages`. It should like something work `https://{subdomain}.ngrok.io/api/messages`.

- Click on the `Bots` menu item from the toolkit and select the bot you are using for this project.  Update the messaging endpoint and press enter to save the value in the Bot Framework.

- Ensure that you've [enabled the Teams Channel](https://docs.microsoft.com/en-us/azure/bot-service/channel-connect-teams?view=azure-bot-service-4.0)

**Configure Bot**

Create a `.env` File in the root directory and add the following lines

- BotId=your-bot-id

- BotPassword=your-bot-password




## Build and run

### `npm install`

### `npm start`

## Deploy to Teams
Start debugging the project by hitting the `F5` key or click the debug icon in Visual Studio Code and click the `Start Debugging` green arrow button.
Alternatively, you can download the the App Manifest from the App Studio tab within the MS Teams Extension for VSCode and then upload the zip file as a custom app to Teams.


