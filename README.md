# Watch History
  <p align="center">Make Sense of Your YouTube Watched History</p>
  <p align="center"><img src="./docs/watch-history-03-done.png"></p>

## Motivation

**User-focused -** It’s your data, and you should be empowered to analyze it! 📊

## Features

**Converts YouTube’s Watch History into a Microsoft Excel Spreadsheet:** Easily organize and explore your viewing habits.

**Cross-Platform Compatibility:** The generated spreadsheet can be opened in Excel, LibreOffice, and ONLYOFFICE.

## How to Use
  Visit Google Takeout and export everything from YouTube.
  
  Run the program with your exported data to transform your watch history into an insightful spreadsheet.

## Planned Future
We’re committed to enhancing your experience:

  **Improved Error Handling:** We’ll handle bad zip, HTML, or JSON files gracefully.

  **Interactive Charts:** Visualize trends and patterns in your viewing history.

  **User-Driven Features:** We welcome your suggestions for additional features!

Remember: You can do it! Take control of your data and explore your YouTube journey.

## Notes From the Author

I was intending to use YouTube's API for the data, but Google [removed the API access](https://developers.google.com/youtube/v3/revision_history#august-11,-2016) to YouTube Watch History in 2016. Therefore, the program processes the Watch History file you can get from Google Takeout instead. The code handles the Takeout ZIP file as well as the "Watch History" JSON and the HTML versions. This program is in early development, with documentation and better error handling coming soon.
