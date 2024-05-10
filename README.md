# Microsoft Teams Migrator

Python app to migrate content of [Microsoft Teams](https://www.microsoft.com/en-us/microsoft-teams/group-chat-software) using the [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview) such as teams, channels, messages, replies, reactions, attachments, etc. in a best effort fashion **for small companies**. For details, see the official [Microsoft Documentation](https://learn.microsoft.com/en-us/microsoftteams/platform/graph-api/import-messages/import-external-messages-to-teams).

## Disclaimer

- Microsoft doesn't recommend using this API for data migration. It does not have the throughput necessary for a typical migration. Decide based on the [service limits](https://learn.microsoft.com/en-us/graph/throttling-limits#microsoft-teams-service-limits) if this app can work at your scale. Otherwise consider 3rd party solutions.
- This app was implemented in a best effort fashion and covers all aspects which were relevant for our migration. You can migrate a channel and compare the outcome with the original and decide if this works for you. Otherwise, feel free to modify the code just for yourself or create a Merge Request on this repo.

## Prerequisits

- User accounts must be present in the new Teams / Tenant ahead of the migration. Otherwise, their messages will be posted in the name of the default account
- Create an [Entra ID app registration](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app) in the old Tenant with the following configuration (as we were not allowed to get application permissions on the old tenant, we configured it this way. If you are allowed, configure it as mentioned below)
  - To authenticate a user from the app, configure the 'Mobile and desktop applications' authentication
  - Grant the app registration the following delegated permissions on the Graph API: Channel.ReadBasic.All, ChannelMember.Read.All, ChannelMessage.Read.All, Group.Read.All, User.Read
- Create an [Entra ID app registration](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app) in the new Tenant with the following configuration
  - Use the OAuth2 client credential flow. For that you need the clientId and a clientSecret
  - Grant the app registration the following application permissions on the Graph API: Channel.Create, Channel.ReadBasic.All, ChannelMember.Read.All, ChannelMessage.Read.All, Group.Read.All, TeamMember.ReadWrite.All, Teamwork.Migrate.All
- The user used to export / read the content of the old Teams MUST have at least read permission on all Teams that you want to export
- Add both app registrations to this python app, see config-old-teams.cfg and config-new-teams.cfg
- Store ClientSecret as environment variable named 'CLIENT_SECRET'
- Configure the default_user
- Configure user_map
- Configure sharepoint_map
- Configure teams_to_export
- Configure channels_to_export

## Setup environment

1. Create venv: python3 -m venv venv
2. Activate venv: source venv/bin/activate
3. Update pip: python -m pip install --upgrade pip
4. Install dependencies with pip for this project: pip install -r requirements.txt

Run / Debug application: python3 main.py

[Documentation](https://code.visualstudio.com/docs/python/debugging) on how to debug Python3 apps in VSCode.

## Limitations

- The General channel of a Team MUST be migrated last because as soon as the General channel's migration is completed, the team's migration will be completed as well. The team's state cannot be put back into migration mode so no other channels and messages can then be imported.
- Reactions cannot be created using app permissions, so this python app adds a section at the end of the message with the information who has reacted with which reaction.
- Messages from users who do not exist in the new teams will be created in the name of a default user and the message contains at the beginning the information who originally posted it.
- The link to attachments stored in the Team's SharePoint will be corrected based on the configurable SharePoint mapping. The data migration from the old SharePoint to the new SharePoint is out of scope and must be done manually. M365 Documents (Word, Excel, etc.) cannot be editet in the new teams after migration, but can when using Sharepoint. If this is an issue, please edit the message manually and link the M365 document from the new SharePoint
- Currently the python app only migrates one Team at a time. One could easily improve it by adding a for loop in the main.py file

## Working with Dev Containers

Use the following extension: https://github.com/microsoft/vscode-dev-containers/tree/main/containers/python-3
