# react-page-feedback

## Summary

A simple SPFx feedback webpart which sends notifications to a Microsoft Teams channel when any user gives a feedback for a page in the portal.
Uses adaptive cards for both the feedback form as well as the [Incoming Webhool Url](https://docs.microsoft.com/en-us/microsoftteams/platform/webhooks-and-connectors/how-to/add-incoming-webhook?WT.mc_id=m365-11878-rwilliams) to post notification message in Teams.

![main-image](./assets/main.png)

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.11-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant?WT.mc_id=m365-11878-rwilliams)
- [Microsoft Graph]()
- [Adaptive cards](adaptivecards.io)
- [Incoming webhooks](https://docs.microsoft.com/en-us/microsoftteams/platform/webhooks-and-connectors/how-to/add-incoming-webhook?WT.mc_id=m365-11878-rwilliams)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](https://developer.microsoft.com/en-us/microsoft-365/dev-program?WT.mc_id=m365-11878-rwilliams)


## Solution

Solution|Author(s)
--------|---------
react-page-feedback.sppkg | [Rabia Williams](https://twitter.com/williamsrabia)

## Version history

Version|Date|Comments
-------|----|--------
1.1|March 10, 2021|Update comment
1.0|January 29, 2021|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

  > Microsoft Graph API  Permission has to be granted at a minimum of User.Read

## Features

- The webhook URL is configurable for ease of change (should ideally be a secret, but used as a feature here for demo purpose)
- Uses Graph to get user information to be clear on the feedback card in Teams conversation.
- Used adaptive cards for easy development in SPFx

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant?WT.mc_id=m365-11878-rwilliams)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview?WT.mc_id=m365-11878-rwilliams)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis?WT.mc_id=m365-11878-rwilliams)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview?WT.mc_id=m365-11878-rwilliams)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
