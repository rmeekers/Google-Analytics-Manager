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

* a Google sheet
* access to at least one Google Analytics account

## Limitations

Write operations to the Google Analytics Management API are still in a limited beta available as developer preview. You can request access to the beta program by filling [this form](https://docs.google.com/a/procurios.eu/forms/d/e/1FAIpQLSf01NWo9R-SOHLKDUH0U4gWHNDBIY-gEI-zqBMG1Hyh3_hHZw/viewform){:target="_blank"}.

# Contact
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
