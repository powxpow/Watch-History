# Watch History
  <p align="center">Make Sense of Your YouTube Watched History</p>
  <p align="center"><img src="./docs/watch-history-03-done.png"></p>

## Motivation

**User-focused -** It’s your data, and you should be empowered to analyze it! 📊

## Features

**Converts YouTube’s Watch History into a Microsoft Excel Spreadsheet:** Easily organize and explore your viewing habits.

**Cross-Platform Compatibility:** Written in Python, the program runs on Windows, Mac, and Linux. The generated spreadsheet can be opened in Excel, LibreOffice, ONLYOFFICE and Gnumeric.

**See Your Data:**
See the videos and how many times you watched them. The links are clickable:
  <p align="center"><img src="./docs/spreadsheet-videos.png"></p>

NOTE: In the test sample above, it looks like "Music Artist" is duplicated, but that is because video link is different! One link points to "music.youtube.com" and the other points to "www.youtube.com". So you can count how many times you viewed a video versus listened to it. This is based on real data I have tested.

See your monthly views as a table and a chart:
  <p align="center"><img src="./docs/spreadsheet-monthly.png"></p>

Your data will show up on multiple tabs:
  <p align="center"><img src="./docs/spreadsheet-tabs.png"></p>

## How to Use
Download the code. Alternatively, you can download the release for Windows or Linux on the right of the screen and unzip it to where you want.

Visit <a href="https://takeout.google.com">Google Takeout</a> and export everything from YouTube.
  
Run the program with your exported data to transform your watch history into an insightful spreadsheet.

Remember: You can do it! Take control of your data and explore your YouTube journey.

## Notes From the Author

I was intending to use YouTube's API for the data, but Google [removed the API access](https://developers.google.com/youtube/v3/revision_history#august-11,-2016) to YouTube Watch History in 2016. Therefore, the program processes the Watch History file you can get from Google Takeout instead. The code can process the Takeout ZIP file as well as the "Watch History" JSON and the HTML versions.
