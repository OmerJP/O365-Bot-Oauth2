# Building your personal Office 365 assistant #

## Summary ##

This sample demonstrates how to use the oauth2 in order to perform the authentication again the Azure Active Directory and grabbing an access token that can be used to consume the Microsoft Graph API for your Office 365 tenant.

[Blog post here https://delucagiuliano.com/building-your-personal-office-365-assistant](https://delucagiuliano.com/building-your-personal-office-365-assistant)

### When to use this pattern? ###
This sample is suitable when you want to implement a personal Office 365 assistant using the context of the user. 


## Applies to

* [Office 365 tenant](https://dev.office.com/sharepoint/docs/spfx/set-up-your-development-environment)

## Solution

Solution|Author(s)
--------|---------
DebbyBot | Giuliano De Luca (MVP Office Development) - Twitter @giuleon

## Version history

Version|Date|Comments
-------|----|--------
1.0 | May 01, 2018 | Initial release

## Disclaimer
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository and follow the instructions below

## Prerequisites ##
 
### 1- Setup the Azure AD application ###

The Bot makes use of Microsoft Graph API (App Only), you need to register a new app in the Azure Active Directory behind your Office 365 tenant using the Azure portal.
