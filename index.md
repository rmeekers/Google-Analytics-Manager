---
layout: default
title: Google Analytics Manager
---
# Background
At one of the projects I've worked on in the past, I had the challenge to manage over a hundred websites on a single platform. All of these websites should be equally tracked in Google Analytics. Each website had his own Google Analytics Property with a minimum of three views and three custom dimensions. So at the end I had to manage a minimum of:

* 5 accounts
* 100 properties
* 300 custom dimensions
* 300 views

No matter if you have to manage five or a hundred properties, keeping them aligned configuration wise is a challenge. That's why I started automating the creation and maintenance of this stuff.

# What's the tool about?
Google Analytics Manager is a Google Apps Script living in a Google sheet.
It enables you to:

* List items:
   * Properties
   * Views
   * Custom Dimensions
   * Filter Links
* Create / Update items:
   * Properties
   * Views
   * Custom Dimensions

# Installation

## Requirements

You'll need the following:

* a Google Sheet – follow the steps below or create a copy from [this sheet](https://docs.google.com/spreadsheets/d/1QrJzquookyryTLuKEI2BoX822jAVcSlg9ExqVw1Eywc/copy)
* access to at least one Google Analytics account

## Installation Steps

### Installation on an existing or blank sheet

1. Open the Sheet and go to *Tools > Script editor*
2. Copy paste script files (*code.gs* and *index.html*)
3. Go to *Resources > Advanced Google Services* and enable the *Google Analytics API*
4. Go to *Resources > Developers Console Project*
5. Enable the *Google Analytics* API
6. Close the Script editor and reload the Sheet
7. Click *Audit GA* in the menu
8. Accept permissions
9. You're set! Create you're first report by following the steps in the sidebar.

### Installation when using the copied sheet

1. Open the Sheet and go to *Tools > Script editor*
2. Go to *Resources > Developers Console Project*
3. Enable the *Google Analytics* API
4. Close the Script editor and reload the Sheet
5. Click *Audit GA* in the menu
6. Accept permissions
7. You're set! Create you're first report by following the steps in the sidebar.

## Limitations

Be aware that write operations to the Google Analytics Management API are still in a limited beta available as developer preview. You can request access to the beta program by filling [this form](https://docs.google.com/a/procurios.eu/forms/d/e/1FAIpQLSf01NWo9R-SOHLKDUH0U4gWHNDBIY-gEI-zqBMG1Hyh3_hHZw/viewform){:target="_blank"}.

# How to use the tool

## Retrieve data from Google Analytics

To retrieve data from Google Analytics you just have to click *GA Manager > Audit GA* in the menu. A sidebar will appear where you can select the account(s) and the type of report you want to generate.
Just click *Audit account(s)* and the report will be inserted on a new sheet. If you regenerate a report of the same type, the data in the existing sheet will be overridden.

## Write data back to Google Analytics

To write data back to Google Analytics, you should first audit an account for the data you want to write back. Just to make sure the sheet is present with the correct columns and so on.
Once the sheet is there, you can either update existing rows or insert new rows. The only thing you have to do it mark a row for inclusion in the first column and then click on *GA Manager > Insert / Update data from the active sheet to GA*.
If there is something wrong, it will tell you.

# Defects and Change Requests

Please file an issue on the project page on GitHub and provide as much information as you have. Please include the steps to reproduce in case of a defect.

# Changelog

See the [Milestones page](https://github.com/rmeekers/Google-Analytics-Manager/milestones?state=closed) on GitHub.

# Contact

You can contact Rutger via:

<a href="https://twitter.com/rutgermeekers" target="_blank" class="icon-link">
    <svg class="icon icon-twitter">
        <use xlink:href="#icon-twitter" />
    </svg>
</a>
<a href="https://github.com/rmeekers" target="_blank" class="icon-link">
    <svg class="icon icon-github">
        <use xlink:href="#icon-github" />
    </svg>
</a>
<a href="https://linkedin.com/in/rutgermeekers" target="_blank" class="icon-link">
    <svg class="icon icon-linkedin">
        <use xlink:href="#icon-linkedin" />
    </svg>
</a>

# License
Licensed under GNU LESSER GENERAL PUBLIC LICENSE Version 3, 29 June 2007.
