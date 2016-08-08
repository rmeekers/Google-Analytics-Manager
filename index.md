---
layout: default
title: Google Analytics Manager
---

## Background
At one of the projects I've worked on in the past, I had the challenge to manage over a hundred websites on a single platform. All of these websites should be equally tracked in Google Analytics. Each website had his own GA Property with a minimum of three views and three custom dimensions. So at the end I had to manage a minimum of:

* 5 accounts
* 100 properties
* 300 custom dimensions
* 300 views

No matter if you have to manage five or a hundred properties, keeping them aligned configuration wise is a challenge. That's why I started automating the creation and maintenance of this stuff.

## What's the tool about?
With Google Analytics Manager you can:

* List items:
   * Properties
   * Views
   * Custom Dimensions
   * Filter Links
* Create / Update items:
   * Properties
   * Views
   * Custom Dimensions

And this is done from within a Google Sheet.

The tool itself is actually a Google Apps Script that lives in a Google Sheet.
