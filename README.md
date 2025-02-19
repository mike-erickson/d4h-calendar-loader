# d4h-calendar-loader
Uploads new activities (events or exercises) from a Google Spreadsheet to your D4H.com team calendar

## How to install 
1. Download the spreadsheet file **D4H Calendar Loader.xlsx**
2. Download the App Script file **D4H Calendar Loader Script.gs**
3. Open the spreadsheet in Google Spreadsheet.  (The script does not work with Microsoft Excel)
4. Click the **Extensions/Apps script** menu in the spreadsheet
5. Copy and paste the entire contents of the App Script file into the **Code.gs** page of the code editor
6. In the first 8 lines of code, uncomment the D4H server domain you use, and comment out all other domains
7. Click the **Save icon** in the App Script editor and close the App Script window
8. Close the spreadsheet file.

## How to set up
1. Re-open the spreadsheet file in Google Spreadsheet
2. Review and accept the permissions required by the script
3. Select the menu item **D4H/Setup API Connection** (you only need to run this step once)
4. Create a _Personal Access Token_ for your D4H account.  Copy and paste the token (a block of text several hundred characters long) when prompted by the spreadsheet
5. Select the menu item **D4H/Test API Connection** to make sure the token and API connection are working
6. If you are using _Location Bookmarks_, select the **D4H/Load Loaction IDs** menu item
7. If you are using _Tags_ in your D4H activities, select the **D4H/Load Tags** menu item
8. Enter the data you want to upload, into the _Calendar Data_ tab of the Google Spreadsheet.  See the _How To Enter Data_ section, below
9. Select the **D4H/Verify Data** menu item.  Fix any issues it reports
10. Select the **D4H/Upload Data** menu item
11. Watch the App Script console and the _Upload Results_ tab for any problems during the upload
12. Log in to D4H and add any remaining data.  The D4H API does not yet support uploading of custom fields
13. When you are done using the script, select the menu item **D4H/Clear D4H connection data** to delete your token from the script's persistent storage.

## How to enter data
1. Review the example data provided.  Then delete it or overwrite with your data
2. Copy all of the formulas to any new rows that you add to the spreadsheet
3. Data in the bright green columns (Type, Title, startsAt, endsAt) is **required**
4. Data in the light green columns (Description, Preplan, BookmarkID, Tag1..Tag6) is **optional**
5. Data in the light red columns (Date, Begin, End, Day, Location, Tag1 name ... Tag 6 name) is **for convenience only** and is not uploaded to D4H
6. **startsAt** and **endsAt** cells must be in ISO8601 format with GMT timezone.  Data in these cells are generated from the data you enter in the **Date**, **Begin**, and **End** cells
7. To enter a multi-day activity, you will have to manually enter the **startsAt** and **endsAt** cells (in ISO8601 format) (overwriting the formulas in those cells)
8. **Description** and **Preplan** cells support some limited HTML formatting.  You have to enter the HTML tags by hand.  Use \<p> as a newline (not \n)
9. If you are using **Location Bookmark IDs**
  * Enter the ID numbers of the bookmarks, in the **BookmarkID** column.
  * The _Location_ column shows a human-readable version of the ID number, based on the bookmarks you downloaded from D4H
  * The script does not yet support directly adding addresses of activities.  Only location bookmarks
  * If you modify the location bookmarks in your D4H account, re-run the **D4H/Load Location IDs** menu function
10. If you are using **Tags** in your activities
  * Enter the Tag numbers into columns _Tag1_ .. _Tag6_
  * The _Tag1 name_ ... _Tag6 name_  columns provide a human readable version of the tags that will be uploaded.
  * If you modify the tags list in your D4H account, re-run the **D4H/Load Tags** menu function

## Tips:
* Rows can be in any order.  Activities can be ordered chronologically, or grouped by activity type, etc.  Blank rows are ignored
* You can re-order the columns, or add new columns.  But do not change the names of the existing columns
* There is no undo feature.  Make sure your data is correct before you upload it!
* Build your master list of activities in a separate spreadsheet.  Then copy in rows of data to the script-enabled spreadsheet for upload
* Start by uploading a few activities at a time.  Review the results in D4H to make sure everying is working the way you want before uploading a large block
* Create a separate Personal Access Token for uploading calendar data.  Give the token a short expiration date, or delete it when you are done uploading
* The account you use to create the Personal Access Token, must have permission to create activities in your D4H team account


## TODO
* Error detection and reporting in this script are not very robust
* Custom fields cannot be uploaded using the API
* D4H API v3 is a work in progress.  Watch for new capabilities from D4H.
* This script does not yet support uploading addresses directly, it only uploads bookmark location IDs
* Abstracting the D4H API functions into a library would be a good idea
* It might be useful to download activities from D4H into the spreadsheet (e.g. Download last year's training calendar.  Modify.  Upload next year's training calendar.)
