# office365-restore-recyclebin
Restore items from recycle bin in SharePoint Online

## Table of Contents
- [office365-restore-recyclebin](#office365-restore-recyclebin)
  - [Table of Contents](#table-of-contents)
  - [Getting started](#getting-started)
    - [Prerequisites](#prerequisites)
    - [Installing](#installing)
      - [PowerShell module: SharePoint Online](#powershell-module-sharepoint-online)
      - [PowerShell module: PnP Online](#powershell-module-pnp-online)
    - [Configuration](#configuration)
      - [Git](#git)
        - [Identity](#identity)
  - [Scripts](#scripts)
    - [Restore-SPORecycleBin](#restore-sporecyclebin)
  - [Release History](#release-history)
  - [Versioning](#versioning)
  - [Authors](#authors)
  - [Articles](#articles)
  - [Tools](#tools)

## Getting started

### Prerequisites

- Git
- PowerShell
- PowerShell module: Microsoft.Online.SharePoint.PowerShell
- PowerShell module: SharePointPnPPowerShellOnline
- Microsoft 365 tenant
- Account with *SharePoint Administrator* privileges

### Installing

#### PowerShell module: SharePoint Online

If you are using PowerShell 5.0 or newer, you can install the SharePoint Management Shell from the PowerShell Gallery by using the following command:
``` powershell
Install-Module -Name Microsoft.Online.SharePoint.PowerShell
Import-Module -Name Microsoft.Online.SharePoint.PowerShell
```

#### PowerShell module: PnP Online

If you are using PowerShell 5.0 or newer, you can install the PnP PowerShell module for SharePoint Online from the PowerShell Gallery by using the following command:
``` powershell
Install-Module -Name SharePointPnPPowerShellOnline
```

### Configuration

#### Git

You can check if your configuration is already done:
``` bash
$ git config --list
```

##### Identity

Git is writing the name and the email address into each commit command. For this you have to configure these information once in your git config file:
``` bash
$ git config --global user.name "John Doe"
$ git config --global user.email johndoe@example.com
```

## Scripts

### Restore-SPORecycleBin

This script was build because of a misconfiguration of a deletion policy and a lot of items was deleted by a task job. The items must be restored from the recycle bin of each SharePoint Online site. A CSV file must be provided with all effected sites to check the recycle bin. With the parameter *RestoreFiles* you can define to get as an output only the identified items or you can also trigger the restore of these items.

``` powershell
  Restore-SPORecycleBin -Organisation <Org Name> -MSTeamsSitesCSV <Path of CSV file with all SPO sites to check> -CSVReport <Path of the report> -RestoreFiles $true
```

## Release History

Please read [release-notes.md](./release-notes.md) for details on getting them.

## Versioning

I use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/gerickes/office365-restore-recyclebin/tags).

## Authors

- Stefan Gericke - *Initial work* - stefan@gericke.name

## Articles

- Microsoft Docs: [Get-PnPRecycleBinItem](https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/get-pnprecyclebinitem?view=sharepoint-ps)
- Microsoft Docs: [Restore-PnpRecycleBinItem](https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/restore-pnprecyclebinitem?view=sharepoint-ps)

## Tools

- [Visual Studio Code](https://code.visualstudio.com/)