# OutlookSetup

## General Usage

A script to autoconfigure Outlook to provide a smoother and more seamless end-user experience in migration scenarios. Once run, the script will do the following:

* Elevates run permissions level to admin
* Shuts down Outlook
* Creates a profile called "MyEmail"
* Creates an HKEY value to set "MyEmail" as the default mail profile
* Reopens Outlook and loads the login screen

## Compatibility

OutlookSetup is compatible with the following versions of Outlook:

* Outlook 2013
* Outlook 2016
* Outlook 2019

## Prerequisites

* Requires admin permissions to run (self elevating)
* Windows 10 or Windows 7 (minimum).

## User interaction requirements

In order to make use of this file, end users will need to dot he following:

* Click to run the file
* Enter a valid email address and password
* Follow next prompts

## Provisioning options

OutlookSetup can be dployed in the following scenarios:

* End user acutated (click to run)
* GPO deployment where users are tied in to AD (active directory)
* Deployment via remote management software
