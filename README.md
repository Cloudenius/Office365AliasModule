# Introduction

This PowerShell Module creates mail aliases in Office 365. These mail aliases are created per domain name or organization. This is to make sure that organizations get unique email addresses.

This module is tested in Azure PowerShell. The author recommends to run the code from Azure PowerShell to simplify authentication to Office 365.

[![PowerShell Gallery](https://img.shields.io/powershellgallery/v/Office365MailAliases.svg?style=flat-square&label=PowerShell%20Gallery)](https://www.powershellgallery.com/packages/Office365MailAliases/)
[![Release](https://github.com/DevSecNinja/Office365AliasModule/actions/workflows/main.yml/badge.svg?branch=master)](https://github.com/DevSecNinja/Office365AliasModule/actions/workflows/main.yml)
[![Lint Code Base](https://github.com/DevSecNinja/Office365AliasModule/actions/workflows/linter.yml/badge.svg?branch=master)](https://github.com/DevSecNinja/Office365AliasModule/actions/workflows/linter.yml)

## Download

The module can be downloaded from [the PowerShell Gallery](https://www.powershellgallery.com/packages/Office365MailAliases) by running the following command in PowerShell:

``` powershell
Install-Module -Name Office365MailAliases
```

## Feature requests

Please create a GitHub issue.

## Build and Test

The code is checked by running the PSScriptAnalyzer extension during the build. Unit tests might follow.

## Contribute

Feel free to open up a PR!
