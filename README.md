# Introduction 
This PowerShell Module creates mail aliases in Office 365. These mail aliases are created per domain name or organization. This is to make sure that organizations get unique email addresses.

This module is tested in Azure PowerShell. The author recommends to run the code from Azure PowerShell to simplify authentication to Office 365.

[![Build Status](https://cloudenius.visualstudio.com/Office%20365%20Alias%20Module/_apis/build/status/Office%20365%20Alias%20Module-CI?branchName=master)](https://cloudenius.visualstudio.com/Office%20365%20Alias%20Module/_build/latest?definitionId=7&branchName=master)
[![Release Status](https://cloudenius.vsrm.visualstudio.com/_apis/public/Release/badge/e1f84d2c-10aa-42f5-b85a-8925cff41305/1/1)](https://cloudenius.visualstudio.com/Office%20365%20Alias%20Module/_build/latest?definitionId=7&branchName=master)
[![CI](https://github.com/Cloudenius/Office365AliasModule/workflows/CI/badge.svg)](https://github.com/Cloudenius/Office365AliasModule/actions?query=workflow%3ACI)

# Download
The module can be downloaded from [the PowerShell Gallery](https://www.powershellgallery.com/packages/Office365MailAliases) by running the following command in PowerShell:
```
Install-Module -Name Office365MailAliases
```

# Feature requests
Open the Work Items [on the Board](https://dev.azure.com/Cloudenius/Office%20365%20Alias%20Module/_workitems/recentlyupdated/).

# Build and Test
The code is checked by running the PSScriptAnalyzer extension during the build. Unit tests might follow.

# Contribute
Feel free to open up a PR!
