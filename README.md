<!--
#-------------------------------------------------------------------------------
# Copyright (C) 2017 EFE.
# All rights reserved. This program and the accompanying materials
# are made available under the terms of the Eclipse Public License v1.0
# which accompanies this distribution, and is available at
# http://www.eclipse.org/legal/epl-v10.html
# 
# Contributors:
#     EFE - initial API and implementation
#-------------------------------------------------------------------------------
-->

# TimesheetLog

This code extracts Timesheet logs from Google Calendars

## Presentation
Inspired by the G Suite [TimeSheet](https://chrome.google.com/webstore/detail/timesheet/nebngbhpfeiihkkpkgfignmdfikfpclb?utm_source=permalink "TimeSheet") add-on.

Unfortunately this add-on is not Open Source, so I rewrote it to implement the feature I wanted:

A simple log report of all the events organised in 8 columns:
* Owner: Calendar name
* Project: Hashtag prefix used in the events
* Title: The text following the hashtag
* Start: Start date of the event
* End date of the event
* Duration: Duration of the event
* Quarter: I also needed the quarter of event
* Month: and the mont hof the event to manage some pivot tables

## Install

For now I don't now out to easily install the code in your speadsheet.
So, you will have to:
* Open your spreadsheet
* From the Tools menu, open the Scipt editor
* Copy/paste the Code.gs content into your Code.gs file
* Create a JavaScript.html file
* Copy/paste the JavaScript.html content into your JavaScript.html file
* Create a Sidebar.html file
* Copy/paste the Sidebar.html content into your Sidebar.html file
* Reload your spreadsheet
* You should find in the Add-ons menu the EFE > Extract Timesheets Sidebar
* Select the calendars you ould like to parse
* Select the Period you would like to extract
* The extractor will create one sheet per calendar and insert the corresponding logs into each one.

## Next plans
* Automate the install process by creating a real G Suite Add-on
* Add some checks like: no calendar selected, wrong period.
* All new contribution is welcome


