# Search Based Message Extension with Single Sign On, Microsoft Graph and SharePoint Online

This sample project demonstrates how to build a search based message extension that a user can use to query data in a SharePoint List and inject that data into a message as a card.

The message extension can be used across Microsoft 365:

- Microsoft Teams
- Microsoft Outlook
- Microsoft 365 Copilot

## Prerequisites

- Teams Toolkit for Visual Studio Code v5.2.0
- Microsoft 365 tenant with [custom apps](https://support.microsoft.com/office/add-an-index-to-a-list-or-library-column-f3f00554-b7dc-44d1-a2ed-d477eac463b0) enabled.
- An active Azure subscription, this sample uses Azure Bot Service for local development

## Get started

### Configure SharePoint site

1. Provision a new SharePoint Online team site using [Product Support](https://lookbook.microsoft.com/details/81e2fee3-02a0-427b-af8b-8c7f42010fde) template from [SharePoint look book](https://lookbook.microsoft.com/).
1. [Create a new Indexed column](https://support.microsoft.com/office/add-an-index-to-a-list-or-library-column-f3f00554-b7dc-44d1-a2ed-d477eac463b0) on the Product list. Set `Title` field as the primary index for the column.

### Configure environment variables

1. Create `env/.env.local` file.

```env
TEAMSFX_ENV=local

OAUTH_CONNECTION_NAME=MicrosoftGraph

BOT_ID=
TEAMS_APP_ID=
BOT_DOMAIN=
BOT_ENDPOINT=
TEAMS_APP_TENANT_ID=
AAD_APP_OBJECT_ID=
AAD_APP_TENANT_ID=
AAD_APP_OAUTH_AUTHORITY=
AAD_APP_OAUTH_AUTHORITY_HOST=
AAD_APP_ACCESS_AS_USER_PERMISSION_ID=
M365_TITLE_ID=
M365_APP_ID=
SPO_HOSTNAME=
SPO_SITE_URL=
```

1. Update `SPO_HOSTNAME` variable to be the hostname of you SharePoint Online tenant, e.g. `contoso.sharepoint.com`.
1. Update `SPO_SITE_URL` variable to be the server relative URL to the Product support SharePoint site, e.g. `sites/productmarketing`.

1. Create `env/.env.local.user` file.

```env
SECRET_BOT_PASSWORD=
```

### Debug

Press F5 and follow the instructions.
