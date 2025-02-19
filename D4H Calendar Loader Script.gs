//--------------------------------------
// UNCOMMENT THE LINE FOR YOUR DOMAIN
//--------------------------------------
  var D4H_API_URL_HEADER = "https://api.team-manager.us.d4h.com/v3/";   // US-Global
//var D4H_API_URL_HEADER = "https://api.team-manager.ap.d4h.com/v3/";   // Asia Pacific
//var D4H_API_URL_HEADER = "https://api.team-manager.ca.d4h.com/v3/";   // Canada
//var D4H_API_URL_HEADER = "https://api.team-manager.eu.d4h.com/v3/";   // Europe


//===============================================================================
// D4H Calendar Loader Script.gs
//
//  Version   3.0
//  Date      18 Feb 2025
//  Author    mErickson@LarimerCountySAR.com
//  Language  Google Apps Script
//  Project   https://github.com/mike-erickson/d4h-calendar-loader 
//
// This script is attached to a Google Spreadsheet, and used to upload data from
// the spreadsheet to a D4H Calendar (d4h.com)
//================================================================================
/* Copyright (c) 2025  Mike Erickson

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
*/

//-----------------------------------------------------------------------------
//  TO USE SOME OF THE INTERNAL TEST FUNCTIONS IN THIS SCRIPT,
//  INSERT YOUR PERSONAL ACCESS TOKEN (c.f. https://help.d4h.com/article/377-obtaining-an-api-access-key)
//   AND YOUR TEAM-ID:
//  ( NOT REQUIRED FOR NORMAL USE OF THIS SCRIPT )
//  The Personal Access Token is several hundred characters long.  Note, when you paste the token in
//  the editor will but linefeeds after any . characters; remove those line feeds!
//-----------------------------------------------------------------------------
var D4H_API_KEY = "FixmeFIXME123456789FixMe555555FixmeFIXME123456789FixMe555555FixmeFIXME123456789FixMe5555550123456789"
                 +"FixmeFIXME123456789FixMe555555FixmeFIXME123456789FixMe555555FixmeFIXME123456789FixMe5555550123456789"
                 +"FixmeFIXME123456789FixMe555555FixmeFIXME123456789FixMe555555FixmeFIXME123456789FixMe5555550123456789"
                 +"FixmeFIXME123456789FixMe555555FixmeFIXME123456789FixMe555555FixmeFIXME123456789FixMe5555550123456789"
                 +"FixmeFIXME123456789FixMe555555FixmeFIXME123456789FixMe555555FixmeFIXME123456789FixMe5555550123456789"
                 +"FixmeFIXME123456789FixMe555555FixmeFIXME123456789FixMe555555FixmeFIXME123456789FixMe5555550123456789";

var D4H_API_TEAM = 0000;     // your team ID

//======================================================
// Names of sheets (tabs) within the Google Spreadsheet
//======================================================
const SOURCEDATA_SHEET_NAME = 'Calendar Data'
const VERIFICATION_SHEET_NAME = 'Verification Results'
const UPLOAD_SHEET_NAME = 'Upload Results'
const LOCID_SHEET_NAME = 'Location Bookmark IDs'
const TAGID_SHEET_NAME = 'D4H Tags'

//======================================================
// Names of the columns in the SOURCEDATA_SHEET_NAME
//   sheet.  Only the columns that are uploaded to D4H
//   need to be tracked here.
//======================================================
const TYPE_COLUMN_NAME = 'Type'
const TITLE_COLUMN_NAME = 'Title'
const STARTSAT_COLUMN_NAME = 'startsAt'
const ENDSAT_COLUMN_NAME = 'endsAt'
const DESCRIPTION_COLUMN_NAME = 'Description'
const PREPLAN_COLUMN_NAME = 'Preplan'
const BOOKMARKID_COLUMN_NAME = 'BookmarkID'
const TAG1_COLUMN_NAME = 'Tag1'
const TAG2_COLUMN_NAME = 'Tag2'
const TAG3_COLUMN_NAME = 'Tag3'
const TAG4_COLUMN_NAME = 'Tag4'
const TAG5_COLUMN_NAME = 'Tag5'
const TAG6_COLUMN_NAME = 'Tag6'
const MAX_TAGS = 6

// Default values for these column numbers.
//  run the getColumnHeaders() function to
//  make sure the user didn't move or delete the columns.
var TYPE_COLUMN = 1
var TITLE_COLUMN = 2
var STARTSAT_COLUMN = 3
var ENDSAT_COLUMN = 4
var DESCRIPTION_COLUMN = 5
var PREPLAN_COLUMN = 6 
var BOOKMARKID_COLUMN = 7
var TAG1_COLUMN = 8
var TAG2_COLUMN = 9
var TAG3_COLUMN = 10
var TAG4_COLUMN = 11
var TAG5_COLUMN = 12
var TAG6_COLUMN = 13


//==================================================================================
// onOpen()
//   Automatically gets run when the spreadsheet file is opened.
// Creates the menu of D4H-related actions
// Displays the initial dialog box
//==================================================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('D4H')
      .addItem('Setup D4H API connection','D4HSetup')
      .addItem('Test D4H API connection', 'D4HTestAPI')
      .addItem('Load Location Bookmark IDs from D4H', 'D4HLoadBookmarks')
      .addItem('Load Tags from D4H', 'D4HLoadTags')
      .addItem('Verify data before uploading','D4HVerify')
      .addItem('Upload calendar to D4H','D4HUpload')
      .addItem('Clear D4H connection data','D4HUnSetup')
      .addToUi();
  ui.alert("D4H Calendar Loader", 
        "This Google Spreadsheet, and its connected AppsScript program, uploads "+
        "events and exercises to your D4H Calendar.\n\n"+
        "Read the instructions (in the \'Instructions\' tab at the bottom of the screen) "+
        "thoroughly before using the uploader!\n\n"+
        "THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS \"AS IS\" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.",
        ui.ButtonSet.OK )
}


//===================================================================================
// D4HSetup()
//   Menu item.
//   Used to set or change the D4H API key
//   Does not save the key if it cannot connect to D4H 
//===================================================================================
function D4HSetup() {
  
  var scriptProperties = PropertiesService.getScriptProperties();
  var APIuser = scriptProperties.getProperty('APIuser');
  var APIkey = scriptProperties.getProperty('APIkey');
  console.log(APIkey)

  var ui = SpreadsheetApp.getUi();
  if(APIkey != null) {
    var result = ui.alert(
      "D4H API Key Change",
      "There is already a saved API key for this spreadsheet.  Do you wish to change it?",
      ui.ButtonSet.YES_NO
    );
    if(result == ui.Button.NO) {
      return
    }
  }

  var result = ui.prompt("D4H API Setup", 
    "To create a Personal Access Token for your account, see https://help.d4h.com/article/377-obtaining-an-api-access-key \n\n"+
    "Enter your D4H Personal Access Token:", 
    ui.ButtonSet.OK_CANCEL );
  var button = result.getSelectedButton();
  var APIkey = result.getResponseText();
  if(button !== ui.Button.OK) { return; }

  // Test the API key
  var getTeamID = D4H_get_team(APIkey)
  if(getTeamID.code !== 200) {
    result = ui.alert("D4H API ERROR", 
        "Could not connect to D4H with that key.\n\nDetails : "+getTeamID.teamID,
        ui.ButtonSet.OK)
    return
  }
  var getUser = D4H_get_user(getTeamID.teamID,APIkey)
  if(getUser.code !== 200) {
    result = ui.alert("D4H API ERROR", 
        "Could not connect to D4H with that key.\n\nDetails : "+getUser.user,
        ui.ButtonSet.OK)
    return
  }

  // we can get user and teamID using the API key, but store them for convenience
  scriptProperties.setProperty('APIkey',APIkey)
  scriptProperties.setProperty('APIuser',getUser.user)
  scriptProperties.setProperty('APIteam',getTeamID.teamID)

  ui.alert("D4H Access Confirmed","D4H user:"+getUser.user+
    "\n\nThe API key is now stored in this script's persistent storage.  "+
    "You won't have to do this setup again unless you change the API key.\n"+
    "To delete the key from persistent storage, use the \'Clear D4H connection data\' menu item",
    ui.ButtonSet.OK)
}


//===================================================================================
// D4HUnSetup
//   Menu item.
//   Deletes the API key from the script's persistent storage
//===================================================================================
function D4HUnSetup () {
  var ui = SpreadsheetApp.getUi();
  result=ui.alert("Clear D4H Information",
    "This will erase the API key and associated information from this script's persistent storage.",
    ui.ButtonSet.OK_CANCEL)
  if(result != ui.Button.OK) {return}

  PropertiesService.getScriptProperties().deleteAllProperties();
  ui.alert("Clear D4H Information",
    "D4H API key and information cleared.",
    ui.ButtonSet.OK)
}



//===================================================================================
// D4HTestAPI()
//   Menu item.
//   Tests the saved D4H API key 
//===================================================================================
function D4HTestAPI() {
  var ui = SpreadsheetApp.getUi();
  var scriptProperties = PropertiesService.getScriptProperties();
  var APIuser = scriptProperties.getProperty('APIuser');
  var APIkey  = scriptProperties.getProperty('APIkey');
  var APIteam = scriptProperties.getProperty('APIteam');

  // Test the API key 
  var getUser = D4H_get_user(APIteam,APIkey)
  if(getUser.code !== 200) {
    result = ui.alert("D4H API ERROR", 
        "Could not connect to D4H API.\n\nDetails : "+getUser.user,
        ui.ButtonSet.OK)
    return
  }

  ui.alert("D4H Access Confirmed","D4H user:  "+getUser.user,ui.ButtonSet.OK)
}


//===================================================================================
// D4HLoadBookmarks()
//   Menu item.
//   Loads the list of Location Bookmark IDs from the team D4H site.
//   Saves them to the sheet named 'Location Bookmark IDs' in this spreadsheet, 
//     overwriting any data previously in that sheet.
//   Bookmark IDs are the easiest way to set the location but they must be specified
//     by number.  You can use a VLOOKUP function to display the human-readable 
//     description of the location, to verify the ID number
//      e.g. =IF(ISBLANK(I2),"",VLOOKUP(I2,'Location Bookmark IDs'!$A$2:$B$200,2))
//===================================================================================
function D4HLoadBookmarks() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    "Load Bookmarks",
    "The script will load the latest set of Location Bookmark IDs from D4H,\n"+
     " overwriting any data in the sheet \'"+LOCID_SHEET_NAME+"\'",
    ui.ButtonSet.OK_CANCEL
  )
  if(result==ui.Button.CANCEL) { return }

  result = get_stored_API_info();
  if(result==null) {
    ui.alert(
      "Setup Error",
      "D4H connection information is missing.  Please run \'Setup D4H API connection\' in the D4H Menu",
      ui.ButtonSet.OK)
    return
  }
  var APIkey  = result.key;
  teamID = D4H_get_team(APIkey).teamID

  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName(LOCID_SHEET_NAME);
  if(sh==null) {
    sh=ss.insertSheet(LOCID_SHEET_NAME)
  } else {
    sh.clear()
  }

  let response=D4H_API_get('team/'+teamID+'/location-bookmarks',APIkey)
  if( checkD4Hresponse(response,ui)==true) { return }

  let bookmarks=JSON.parse(response).results
  var data=[];
  for(i=0; i<bookmarks.length; i++) {
    data[i]=[]
    data[i][0] = bookmarks[i].id;
    data[i][1] = bookmarks[i].title;
    data[i][2] = bookmarks[i].address.street
    data[i][3] = bookmarks[i].address.town
    data[i][4] = bookmarks[i].address.region
    data[i][5] = bookmarks[i].address.postcode
    data[i][6] = bookmarks[i].address.country
    data[i][7] = bookmarks[i].location.coordinates[0]
    data[i][8] = bookmarks[i].location.coordinates[1]
  }
  var r=sh.getRange("A1:I1")
  r.setValues([['ID','Title', 'Street','City','State','ZIP','Country','Lat','Lon']])
  r=sh.getRange(2,1,i,9)
  r.setValues(data)
}


//===================================================================================
// D4HLoadTags()
//   Menu item.
//   Loads the list of Tags from the team D4H site.
//   Saves them to the sheet named 'Tags' in this spreadsheet, 
//     overwriting any data previously in that sheet.
//   You can use a VLOOKUP function to display the human-readable 
//     description of the location, to verify the ID number
//      e.g. =IF(ISBLANK(I2),"",VLOOKUP(I2,'Tags'!$A$2:$B$200,2))
//===================================================================================
function D4HLoadTags() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    "Load Tags",
    "The script will load the latest set of Tags from D4H,\n"+
     " overwriting any data in the sheet \'"+TAGID_SHEET_NAME+"\'.",
    ui.ButtonSet.OK_CANCEL
  )
  if(result==ui.Button.CANCEL) { return }

  result = get_stored_API_info();
  if(result==null) {
    var result = ui.alert(
      "Setup Error",
      "D4H connection information is missing.  Please run \'Setup D4H API connection\' in the D4H Menu",
      ui.ButtonSet.OK)
    return
  }
  var APIkey  = result.key;
  teamID = D4H_get_team(APIkey).teamID

  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName('D4H Tags');
  if(sh==null) {
    sh=ss.insertSheet('D4H Tags')
  } else {
    sh.clear()
  }

  let response=D4H_API_get('team/'+teamID+'/tags',APIkey)
  if( checkD4Hresponse(response,ui)==true) { return }

  let tags=JSON.parse(response).results
  var data=[];
  for(i=0; i<tags.length; i++) {
    data[i]=[]
    data[i][0] = tags[i].id;
    data[i][1] = tags[i].title;
  }
  var r=sh.getRange("A1:B1")
  r.setValues([['Tag ID','Title']])
  r=sh.getRange(2,1,i,2)
  r.setValues(data)
}


//===================================================================================
// D4HVerify()
//  Menu item
//  This does a precheck of the data
//===================================================================================
function D4HVerify() {
  var ui = SpreadsheetApp.getUi()
  var result = ui.alert(
    "Data verification",
    "The script will run some verification checks on the data without uploading anything to D4H.\n"+
    "  Results will be placed in the \'Verification Results\' sheet.",
    ui.ButtonSet.OK_CANCEL
  )
  if(result!=ui.Button.OK) { return}

  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var resultsSheet=ss.getSheetByName(VERIFICATION_SHEET_NAME);
  if(resultsSheet==null) {
    resultsSheet=ss.insertSheet(VERIFICATION_SHEET_NAME)
  } else {
    resultsSheet.clear()
  }

  var dataSheet=ss.getSheetByName(SOURCEDATA_SHEET_NAME);
  if(dataSheet==null) {
        result = ui.alert("ERROR", 
          "Could not find a sheet named Calendar Data in this spreadsheet.",
          ui.ButtonSet.OK)
    return
  }
  setColumnIDs()  

  // Load the source data into a 2D array
  SpreadsheetApp.setActiveSheet(resultsSheet);
  var numRows=dataSheet.getLastRow()-1;
  var numCols=dataSheet.getLastColumn()
  var data=dataSheet.getRange(2,1,numRows,numCols).getDisplayValues()

  // Load the LocID data into a 2D array
  var locIDs = loadIndexSheet(ss, LOCID_SHEET_NAME)

  // Load the Tags data into a 2D array
  var tagIDs = loadIndexSheet(ss, 'D4H Tags')


  // Check each row
  SpreadsheetApp.setActiveSheet(resultsSheet);
  var c=resultsSheet.getRange('A1')
  c.setValue('Verification Results  ')
  var outRow = 2;
  for(idx=0;idx<numRows;idx++) {
    let type = data[idx][TYPE_COLUMN-1];
    let title = data[idx][TITLE_COLUMN-1]
    let startsAt = data[idx][STARTSAT_COLUMN-1]
    let endsAt = data[idx][ENDSAT_COLUMN-1]
    let description = data[idx][DESCRIPTION_COLUMN-1]
    let preplan = data[idx][PREPLAN_COLUMN-1]
    let locID = data[idx][BOOKMARKID_COLUMN-1]
    let tags = [ data[idx][TAG1_COLUMN-1],
                 data[idx][TAG2_COLUMN-1],
                 data[idx][TAG3_COLUMN-1],
                 data[idx][TAG4_COLUMN-1],
                 data[idx][TAG5_COLUMN-1],
                 data[idx][TAG6_COLUMN-1] ]

    //Logger.log("%d: [%s] [%s] [%s] [%s] [%s] [%s] [%s]", idx, type, title, startsAt, endsAt, description, preplan, locID)

    let msg="Calendar Data,Row "+outRow
    c=resultsSheet.getRange(outRow,1)
    c.setValue(msg)

    msg=''
    let foundError = false
    let foundWarning = false

    var msg1='OK'
    var msg2=''
    if(type=='' && title=='' && startsAt == '' && endsAt == '' && description == '' && preplan=='' && locID=='') {
      msg2 = "(ignoring blank row)"
    } else {

      if(type==null || type=='') {
        msg1  = "ERROR"
        msg2 += "Type (Event/Exercise) is required (cell A"+outRow+")   "
        foundError=true
      }

      if(type!=='Exercise' && type!=='Event') {
        msg1  = "ERROR"
        msg2 += "Type must be \'Event\' or \'Exercise\' (cell A"+outRow+")   "
        foundError=true
      }

      if(title==null || title=='') {
        msg1  = "ERROR"
        msg2 += "Title is required (cell B"+outRow+")   "
        foundError=true
      }

      if(!isISO8601(startsAt)) {
        msg1 = "ERROR"
        msg2 += "Bad starting date and time (cell C"+outRow+")   "
        foundError=true
      }

      // D4H allows you to define an activity without an end time
      //  But it shows up as 0 minutes length with the end time as '-'
      //  Lets make that illegal for our team.  Better to have a best-guess 
      //  end time than none at all!
      if(endsAt==null || endsAt=='') {
        if(!foundError) { msg1 = 'ERROR' }
        msg2 += "No end date and time (cell D"+outRow+")   "
        foundWarning=true
      } else {
        if(!isISO8601(endsAt)) {
        msg1 = "ERROR"
        msg2 += "Bad ending date and time (cell D"+outRow+")   "
        foundError=true
        }
      }

      if(locID) {
        if(!getID(locID, locIDs)) {
          if(!foundError) { msg1 = 'WARNING' }
          msg2 += "(WARN) BookmarkID "+locID+" not found in the index sheet."
          foundWarning=true
        }
      }

      for(i=0; i<6; i++) {
        if(tags[i]) {
          if(!getID(tags[i], tagIDs)) {
            if(!foundError) { msg1 = 'WARNING' }
            msg2 += "(WARN) Tag ID "+tags[i]+" not found in the index sheet."
            foundWarning=true
          }
        }
      }

      //TODO: make sure Description and PrePlan do not have any illegal characters
      //  \n will be replaced with <p> as part of the upload routine
      /*
      if(/\n/.test(description)) {
        msg1 = "WARN"
        msg2 += "Illegal characters in description text. (cell G"+outRow+")    "
        foundError=true
      }
      if(/\n/.test(preplan)) {
        msg1 = "ERROR"
        msg2 += "Illegal characters in preplan text. (cell H"+outRow+")    "
        foundError=true
      }
      */
  
    }

    // TODO: Combine these into one row write
    c=resultsSheet.getRange(outRow,2)
    c.setValue(msg1)
    c=resultsSheet.getRange(outRow,3)
    c.setValue(msg2)
    resultsSheet.insertRowAfter(outRow)
    outRow++
  }
}


//===================================================================================
// D4HUpload()
//   Menu item
//   This does the upload to D4H
//===================================================================================
function D4HUpload() {
  var ui = SpreadsheetApp.getUi()
  var result = ui.alert(
    "Calendar Upload",
    "The script will upload the calendar data to D4H.\n"+
    "Make sure you run the '\Verify data before uploading\' menu item and fix all reported errors before uploading.  \n\n"+
    "Results of the upload will be placed in the \'"+UPLOAD_SHEET_NAME+"\' sheet.\n\n"+
    "NOTE - there is no undo for this feature.  To change or delete calendar data once it is in D4H, you will have to use the D4H user interface.",
    ui.ButtonSet.OK_CANCEL
  )
  if(result!=ui.Button.OK) { return }

  result = get_stored_API_info();
  if(result==null) {
    var result = ui.alert(
      "Setup Error",
      "D4H connection information is missing.  Please run \'Setup D4H API connection\' in the D4H Menu",
      ui.ButtonSet.OK)
    return
  }
  var APIkey  = result.key;
  var teamID = D4H_get_team(APIkey).teamID

  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var resultsSheet=ss.getSheetByName(UPLOAD_SHEET_NAME);
  if(resultsSheet==null) {
    resultsSheet=ss.insertSheet(UPLOAD_SHEET_NAME)
  } else {
    resultsSheet.clear()
  }

  var dataSheet=ss.getSheetByName(SOURCEDATA_SHEET_NAME);
  if(dataSheet==null) {
        result = ui.alert("ERROR", 
          "Could not find a sheet named Calendar Data in this spreadsheet.",
          ui.ButtonSet.OK)
    return
  }
  setColumnIDs();

  // Load the source data into a 2D array
  SpreadsheetApp.setActiveSheet(resultsSheet);
  var numRows=dataSheet.getLastRow()-1;
  var numCols=dataSheet.getLastColumn()
  var data=dataSheet.getRange(2,1,numRows,numCols).getDisplayValues()

  // Upload each row to D4H
  SpreadsheetApp.setActiveSheet(resultsSheet);
  var c=resultsSheet.getRange('A1')
  c.setValue('Results of data upload: ')
  var outRow = 2;
  for(idx=0;idx<numRows;idx++) {
    let type = data[idx][TYPE_COLUMN-1];
    let title = data[idx][TITLE_COLUMN-1]
    let startsAt = data[idx][STARTSAT_COLUMN-1]
    let endsAt = data[idx][ENDSAT_COLUMN-1]
    let description = data[idx][DESCRIPTION_COLUMN-1]
    let preplan = data[idx][PREPLAN_COLUMN-1]
    let locID = data[idx][BOOKMARKID_COLUMN-1]
    let tags = [ data[idx][TAG1_COLUMN-1],
                 data[idx][TAG2_COLUMN-1],
                 data[idx][TAG3_COLUMN-1],
                 data[idx][TAG4_COLUMN-1],
                 data[idx][TAG5_COLUMN-1],
                 data[idx][TAG6_COLUMN-1] ]

    let msg="Calendar Data,Row "+outRow
    c=resultsSheet.getRange(outRow,1)
    c.setValue(msg)

    // -- create the D4H activity
    var eType;
    if(type=='Event') { eType='events'
    }  else           { eType='exercises' }
    
    var toCalendar = {
      referenceDescription: title,
      startsAt:             startsAt,
      fullTeam:             true
    }
    if(endsAt != '')        {toCalendar.endsAt      = endsAt }
    description.replace('/\n/g','<p>')    //FIXME - this doesnt work
    if(description != '')   {toCalendar.description = description}
    description.replace('/\n/g','<p>')    //FIXME - this doesnt work
    if(preplan != '')       {toCalendar.plan        = preplan}
    if(locID != '')         {toCalendar.locationBookmarkId = locID}

    //FIXME: check for blank row and skip it
    //FIXME: clip title to 50 char len
    //FIXME: clip description, preplan to 100char len

    var result = D4H_API_post('team/'+teamID+'/'+eType, JSON.stringify(toCalendar), APIkey)
    msg1 = checkD4Hresponse(result,ui)
    msg2 = ""
    if(msg1==false) {
      // activity create was successful, continue with creating tags
      var rsp=result.getContentText()
      var jrsp=JSON.parse(rsp)
      var activityID = jrsp.id
      // replace result from checkD4Hresponse with an OK msg
      msg1="OK ("+activityID+")  "+startsAt+"   "+title

      // -- Set the activity's tags
      var strTagArray = '['
      var numTagsFound = 0
      // eliminate any gaps in the tags
      for(i=0;i<MAX_TAGS; i++) {
        if (tags[i]>0) {
          if(numTagsFound>0) {strTagArray += ','}
          strTagArray += tags[i]
          numTagsFound++
        }
      }
      strTagArray += ']'

      if(numTagsFound>0) {
        result = D4H_API_post('team/'+teamID+'/'+eType+'/'+activityID+'/tags', 
                              "{\"tagIds\":"+strTagArray+"}",
                              APIkey)
        msg1 += "   tags:"+strTagArray
        rsp=result.getContentText()
        jrsp=JSON.parse(rsp)
        msg2=JSON.stringify(jrsp)
      }

    }

    // Save the results of the API calls in the Upload Results sheet 
    c=resultsSheet.getRange(outRow,2)
    c.setValue(msg1)
    c=resultsSheet.getRange(outRow,3)
    c.setValue(msg2)

    resultsSheet.insertRowAfter(outRow)
    outRow++
  }
}



//===================================================================================
// UTILITY FUNCTIONS
//===================================================================================

//--------------------------------------------
// Returns the user associated with the API key.
// Returns an object {'code'  = response code.  200 is OK.  429 is rate limited
//                    'user'  = D4H username  (or error code details if .code != 200)}
//--------------------------------------------
function D4H_get_user (teamID, APIkey) {

  let response=D4H_API_get('team/'+teamID+'/whoami',APIkey)
  let code = response.getResponseCode()

  if(code != 200){
    username = 'ERROR '+code+' : '+response.getContentText().toString()
  }
  else {
    jResponse = JSON.parse(response)
    username=jResponse.context.name;
  }
   return {code : response.getResponseCode(),
           user : username}
}
function test_D4H_get_user() {
  response=D4H_get_user(D4H_API_TEAM,D4H_API_KEY)
  Logger.log("code="+response.code+", teamID="+response.user)
  response=D4H_get_user(0,D4H_API_KEY)
  Logger.log("code="+response.code+", teamID="+response.user)
  response=D4H_get_user(D4H_API__KEY,0)
  Logger.log("code="+response.code+", teamID="+response.user)
}


//--------------------------------------------
// Returns the team associated with the API key.
// Useful to test if the API key is working.
// Returns an object {'code'    = response code.  200 is OK.  429 is rate limited
//                    'teamID'  = team ID  (or error code details if .code != 200)}
//--------------------------------------------
function D4H_get_team (APIkey) {
  let response=D4H_API_get('whoami',APIkey)
  let code = response.getResponseCode()
  if(code != 200) {
    teamID = 'ERROR '+code+' : '+response.getContentText().toString()
  }
  else {
    jResponse = JSON.parse(response)
    teamID = jResponse.members[0].owner.id;
  }
  return {code   : code,
          teamID : teamID}
}
function test_D4H_get_team() {
  var response=D4H_get_team(D4H_API_KEY)
  Logger.log("code="+response.code+", teamID="+response.teamID)

  response=D4H_get_team(0)
  Logger.log("code="+response.code+", teamID="+response.teamID)

}


//--------------------------------------------
// Performs an API get to D4H
// Returns the raw data of the response.
//--------------------------------------------
function D4H_API_get(path, APIkey) {
  let config = {
    headers           : { "Authorization" : "Bearer " + APIkey },
    muteHttpExceptions: true,
    method            : 'get'
  }
  let fullAPIURL = D4H_API_URL_HEADER + path;
  //Logger.log(UrlFetchApp.getRequest(fullAPIURL, config))  // preview the fetch
  let data = UrlFetchApp.fetch(fullAPIURL, config);
  return data
}

//--------------------------------------------
// Performs an API POST to D4H
// Returns the raw data of the response.
//--------------------------------------------
function D4H_API_post(path, payload, APIkey) {
    let config = {
      headers: { "Authorization" : "Bearer " + APIkey },
      muteHttpExceptions: true,
      'method'          : 'post',
      'contentType'     : 'application/json',
      'payload'         : payload
    }

    let fullAPIURL = D4H_API_URL_HEADER + path;
    //Logger.log(UrlFetchApp.getRequest(fullAPIURL, config))  // preview the fetch
    let response = UrlFetchApp.fetch(fullAPIURL, config);
    //Logger.log("getResponseCode() "+response.getResponseCode());
    //Logger.log("getContentText() "+response.getContentText().toString());
    return response
}

//--------------------------------------------
// Performs an API PATCH to D4H
// Returns the raw data of the response.
//--------------------------------------------
function D4H_API_patch(path, payload, APIkey) {
    let config = {
      headers: { "Authorization" : "Bearer " + APIkey },
      muteHttpExceptions: true,
      'method'          : 'patch',
      'contentType'     : 'application/json',
      'payload'         : payload
    }

    let fullAPIURL = D4H_API_URL_HEADER + path;
    //Logger.log(UrlFetchApp.getRequest(fullAPIURL, config))  // preview the fetch
    let response = UrlFetchApp.fetch(fullAPIURL, config);
    //Logger.log("getResponseCode() "+response.getResponseCode());
    //Logger.log("getContentText() "+response.getContentText().toString());
    return response
}

//--------------------------------------------
// Checks the return code of the HTTP response
// If an HTTP error is found, it returns the error code
// If 'ui' parameter is provided, it alerts the user with
//  that error.
//  Error 429 = rate limit  TODO: Does D4H send a retry-after header?
//  Error 403 = Forbidden.  I get this if I try to create an activity with a long description.  But I can create one in the UI with the same description.  What is the actual limit that D4H enforces?
//--------------------------------------------
function checkD4Hresponse(response, ui) {
  let code = response.getResponseCode()
  if(code != 200) {
    errtext = 'ERROR '+code+' : '+response.getContentText().toString()
    if(ui) {  ui.alert('ERROR', errtext, ui.ButtonSet.OK) }
    return errtext
  }
  return false
}

//--------------------------------------------
// Loads the stored API info from Properties
//  returns {user, key, team}
//  returns null if Properties not set
//--------------------------------------------
function get_stored_API_info() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var APIuser = scriptProperties.getProperty('APIuser');
  var APIkey  = scriptProperties.getProperty('APIkey');
  var APIteam = scriptProperties.getProperty('APIteam');

  if((APIuser==null) || (APIkey==null) || (APIteam==null)) {
    return null
  }
  return {user:APIuser, key:APIkey, team:APIteam}
}



//-----------------------------------------------
// ISO8601(date) returns a text string of date, converted to ISO8601 format
//  default output is GMT timezone, can be overridden.
// c.f.  ISO8601 format https://en.wikipedia.org/wiki/ISO_8601#Time_offsets_from_UTC
// c.f.  https://developers.google.com/apps-script/reference/utilities/utilities#formatDate(Date,String,String)
//-----------------------------------------------
function ISO8601(date, tz="GMT") {
  if(date<40000) return ''
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var timezone=ss.getSpreadsheetTimeZone()
  var formatted_date = Utilities.formatDate(date, tz, "yyyy-MM-dd'T'HH:mm:ssX")
  //Logger.log(formatted_date)
  return formatted_date
}

function test_ISO8601() {
  var now = new Date()
  Logger.log(ISO8601(now))
  Logger.log(ISO8601(now,"MDT"))
  Logger.log(ISO8601(now, "MST"))
  summer = new Date(2025,6,1,12,0,0);   // June 1 2025 12:00 noon
  Logger.log(ISO8601(summer))
  Logger.log(ISO8601(summer, "MST"))
}

//----------------------------------------------
// isISO8601(date) returns true if date is a properly formatted ISO8601 string
//----------------------------------------------
function isISO8601(isoString) {
  if (!/\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z/.test(isoString)) {
    return false;
  }
  try {
    const testDate = new Date(isoString);
    return !isNaN(testDate.getTime()) /*&& testDate.toISOString() === isoString*/;
  } catch(e) {
    return false;
  }
}

function test_isISO8601() {
  Logger.log('1 '+isISO8601('2025-01-01T07:00:00Z') )
  Logger.log('2 '+isISO8601('2020-01-01T12:34:56Z') )
  Logger.log('3 '+isISO8601('') ) 
  Logger.log('4 '+isISO8601('2020-01-33T00:00:00Z') )
  Logger.log('5 '+isISO8601('0000-01-01T00:00:00Z') )
  Logger.log('6 '+isISO8601('2020-01-01T01:00:00') )
  Logger.log('7 '+isISO8601('2020-14-14T00:00:00Z'))
  Logger.log('8 '+isISO8601('2020-12-14T30:00:00Z'))
}



//---------------------------------------------------
// loads a 2D data array from one of the helper
// sheets of index data from D4H (LocID or Tags)
//---------------------------------------------------
function loadIndexSheet(spreadsheet, sheetName) {
  var sheet=spreadsheet.getSheetByName(sheetName)
  if (sheet) {
    var lastRow=sheet.getLastRow()-1
    var range=sheet.getRange(2,1,lastRow,2)
     return range.getValues()
  }
}



//------------------------------------------------------------------
// getID(ID, table)  Looks up the description associated with the ID
//  table is the data array loaded from the D4H helper sheet
// returns false if ID not found
//------------------------------------------------------------------
function getID(ID, table){
  for(i=0;i<table.length;i++) {
    if (table[i][0]==ID) {return table[i][1]}
  }
  return false
}

function test_getID() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var tags = new Array;
  tags = loadIndexSheet(ss, 'D4H Tags', tags)
  //Logger.log(tags.toString())

  Logger.log('1 '+getID('',tags))
  Logger.log('2 '+getID(0,tags))
  Logger.log('3 '+getID(30535,tags))
  Logger.log('4 '+getID(33439,tags))
  Logger.log('5 '+getID(31000,tags))

  var locIDs = new Array;
  locIDs = loadIndexSheet(ss, 'Location Bookmark IDs', locIDs)
  //Logger.log(locIDs.toString())

  Logger.log('1 '+getID('',locIDs))
  Logger.log('2 '+getID(0,locIDs))
  Logger.log('3 '+getID(4858,locIDs))
  Logger.log('4 '+getID(4827,locIDs))
  Logger.log('5 '+getID(5001,locIDs))

}


//--------------------------------------------------------------
// setColumnIDs  Sets the column index variables (TYPECOL, DESCCOL, etc)
//  by searching row 1 of the spreadsheet for the expected column titles.
//  This lets the script recover in case the user rearranges the columns.
//--------------------------------------------------------------
function setColumnIDs() {
  TYPE_COLUMN        = findHeaderText(TYPE_COLUMN_NAME);
  TITLE_COLUMN       = findHeaderText(TITLE_COLUMN_NAME);
  STARTSAT_COLUMN    = findHeaderText(STARTSAT_COLUMN_NAME);
  ENDSAT_COLUMN      = findHeaderText(ENDSAT_COLUMN_NAME);
  DESCRIPTION_COLUMN = findHeaderText(DESCRIPTION_COLUMN_NAME);
  PREPLAN_COLUMN     = findHeaderText(PREPLAN_COLUMN_NAME);
  BOOKMARKID_COLUMN  = findHeaderText(BOOKMARKID_COLUMN_NAME);
  TAG1_COLUMN        = findHeaderText(TAG1_COLUMN_NAME);
  TAG2_COLUMN        = findHeaderText(TAG2_COLUMN_NAME);
  TAG3_COLUMN        = findHeaderText(TAG3_COLUMN_NAME);
  TAG4_COLUMN        = findHeaderText(TAG4_COLUMN_NAME);
  TAG5_COLUMN        = findHeaderText(TAG5_COLUMN_NAME);
  TAG6_COLUMN        = findHeaderText(TAG6_COLUMN_NAME);
  // return true (1) if all columns got set OK.
  return(TYPE_COLUMN!=null & TITLE_COLUMN!=null & STARTSAT_COLUMN!=null & ENDSAT_COLUMN!=null
         & DESCRIPTION_COLUMN!=null & PREPLAN_COLUMN!=null & BOOKMARKID_COLUMN!=null
         & TAG1_COLUMN!=null & TAG2_COLUMN!=null & TAG3_COLUMN!=null & TAG4_COLUMN!=null 
         & TAG5_COLUMN!=null & TAG6_COLUMN!=null )
}
function testsetColumnIDs() { 
  console.log(setColumnIDs())
  console.log(" TYPE "+TYPE_COLUMN)
  console.log(" TITLE "+TITLE_COLUMN)
  console.log(" STARTSAT "+STARTSAT_COLUMN)
  console.log(" ENDSAT "+ENDSAT_COLUMN)
}


//--------------------------------------------------------------
// Searches Row 1 of the Source Data sheet, for the given text
//  Returns the A1 notation of the first cell where the text is found
//  or null if not found
//--------------------------------------------------------------
function findHeaderText(searchText, sheetName=SOURCEDATA_SHEET_NAME, range="A1:A100") {
  var dataSheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SOURCEDATA_SHEET_NAME);
  if(dataSheet==null) {
        result = ui.alert("ERROR", 
          "Could not find a sheet named Calendar Data in this spreadsheet.",
          ui.ButtonSet.OK)
    return null
  }
  var numCols=dataSheet.getLastColumn()
  var data=dataSheet.getRange(1,1,1,numCols).getDisplayValues()

  for(var c=0; c<numCols; c+=1) {
    if(data[0][c]==searchText) return c+1
  }
  var ui = SpreadsheetApp.getUi();
  ui.alert('Could not find a column named '+searchText+' in Row 1 of the sheet named '+SOURCEDATA_SHEET_NAME+'.\n\n  You may re-order the columns in the spreadsheet but dont delete or rename them!');
  return null
}
//--tester for findHeaderText()
function testfindHeaderText() {
  console.log("1:   "+ findHeaderText("Type") )
  console.log("2:   "+ findHeaderText(TITLE_COLUMN_NAME) )
  console.log("3:   "+ findHeaderText(TAG1_COLUMN_NAME))
  console.log("4: F "+ findHeaderText("updog"))
}

//=============================================================================
// API testing functions
//  Used for development to figure out how the API works
//  You will need to adjust for the IDs in your system, if you want to use these
//=============================================================================
function test_get_exercise(activity_id=972600) {
  let config = {
    headers: {
      "Authorization" : "Bearer " + D4H_API_KEY
    },
    muteHttpExceptions: true,
      "method" : "get"
  }
  let fullAPIURL = D4H_API_URL_HEADER+'team/'+D4H_API_TEAM+'/exercises/'+activity_id
  Logger.log(UrlFetchApp.getRequest(fullAPIURL, config));  // preview the fetch
  let data = UrlFetchApp.fetch(fullAPIURL, config)
  Logger.log(data)

  let response = JSON.parse(data);
  Logger.log(response);
  Logger.log(response.referenceDescription);
  Logger.log(response.startsAt);
  Logger.log(response.endsAt);
  Logger.log(response.resourceType)
}


// NOTE
//   Creating an activity with the locationBookmark does NOT set the address or lat/lon
//   .fullTeam defaults to FALSE.  (is this broken - reverse sense?)
//   \n not allowed in rich text.  use <p> for CR/LF
//   other simple HTML is allowed
function test_create_exercise() {
    // Set details, location, etc using the API
      let payload = "{\"referenceDescription\":\"CalendarLoader Test 16\""+
      ",\"startsAt\":\"2026-01-01T14:03:00Z\""+
      ",\"endsAt\":\"2026-01-01T15:03:00Z\""+
      ",\"description\":\"the description<p> on two lines\""+
      ",\"plan\":\"the <b>plan</b>\""+
      ",\"address\":{\"town\":\"Fort Collins\"}"+
      ",\"locationBookmarkId\":\"4904\""+
      ",\"fullTeam\":false"+
      "}"

    let config = {
      headers: {
        "Authorization" : "Bearer " + D4H_API_KEY
      },
      muteHttpExceptions: true,
      "method" : "post",
      'contentType': 'application/json',
      "payload" : payload
    }

    let fullAPIURL = D4H_API_URL_HEADER+'team/'+D4H_API_TEAM+'/exercises/'

    Logger.log(UrlFetchApp.getRequest(fullAPIURL, config))  // preview the fetch
    let response = UrlFetchApp.fetch(fullAPIURL, config);
    Logger.log("getResponseCode() "+response.getResponseCode());
    Logger.log("getContentText() "+response.getContentText().toString());

//    if(checkD4Hresponse(response)==false) {Logger.log("check-->FAIL")}
//    else {
      jResponse = JSON.parse(response)
      Logger.log("jResponse: "+jResponse)
      Logger.log("refID"+jResponse.reference)
//    }
}


function test_set_tags() {
    let payload ="{\"tagIds\":[30800,30801]}"

    let config = {
      headers: {"Authorization" : "Bearer " + D4H_API_KEY },
      muteHttpExceptions: true,
      "method" : "post",
      'contentType': 'application/json',
      "payload" : payload
    }

    let fullAPIURL = D4H_API_URL_HEADER+'team/'+D4H_API_TEAM+'/exercises/940123/tags'
    Logger.log(UrlFetchApp.getRequest(fullAPIURL, config))  // preview the fetch
    let response = UrlFetchApp.fetch(fullAPIURL, config);
    Logger.log("getResponseCode() "+response.getResponseCode());
    Logger.log("getContentText() "+response.getContentText().toString());

      jResponse = JSON.parse(response)
      Logger.log("jResponse: "+jResponse)
      Logger.log("id"+jResponse.id)   
}


function test_update_exercise() {
    // Change a few details using the API
      let payload = "{"+
      "\"address\":{\"town\":\"Loveland2\"}"+
      ",\"fullTeam\":false"+
      "}"

    let config = {
      headers: {
        "Authorization" : "Bearer " + D4H_API_KEY
      },
      muteHttpExceptions: true,
      "method" : "patch",
      'contentType': 'application/json',
      "payload" : payload
    }

    let fullAPIURL = D4H_API_URL_HEADER+'team/'+D4H_API_TEAM+'/exercises/972600'

    Logger.log(UrlFetchApp.getRequest(fullAPIURL, config))  // preview the fetch
    let response = UrlFetchApp.fetch(fullAPIURL, config);
    Logger.log("getResponseCode() "+response.getResponseCode());
    Logger.log("getContentText() "+response.getContentText().toString());

      jResponse = JSON.parse(response)
      Logger.log("jResponse: "+jResponse)
      Logger.log("refID"+jResponse.reference)

}



function test_get_event() {
  let config = {
    headers: { "Authorization" : "Bearer " + D4H_API_KEY },
    muteHttpExceptions: true,
      "method" : "get"
  }
  let activity_id = 940150; 
  let fullAPIURL = D4H_API_URL_HEADER+'team/'+D4H_API_TEAM+'/events/'+activity_id
  Logger.log(UrlFetchApp.getRequest(fullAPIURL, config));  // preview the fetch
  let data = UrlFetchApp.fetch(fullAPIURL, config)
  Logger.log(data)

  let response = JSON.parse(data);
  Logger.log(response);
  Logger.log(response.referenceDescription);
  Logger.log(response.startsAt);
  Logger.log(response.endsAt);
  Logger.log(response.resourceType)
}


function test_create_event() {

      let payload = "{\"referenceDescription\":\"CalendarLoader Test 16\""+
      ",\"startsAt\":\"2026-01-01T14:03:00Z\""+
      ",\"endsAt\":\"2026-01-01T15:03:00Z\""+
      ",\"description\":\"the description<p> on two lines\""+
      ",\"plan\":\"the <b>plan</b>\""+
      ",\"address\":{\"town\":\"Fort Collins\"}"+
      ",\"locationBookmarkId\":\"4904\""+
      "}"

    let config = {
      headers: {
        "Authorization" : "Bearer " + D4H_API_KEY
      },
      muteHttpExceptions: true,
      "method" : "post",
      'contentType': 'application/json',
      "payload" : payload
    }

    let fullAPIURL = D4H_API_URL_HEADER+'team/'+D4H_API_TEAM+'/events'

    Logger.log(UrlFetchApp.getRequest(fullAPIURL, config))  // preview the fetch
    let response = UrlFetchApp.fetch(fullAPIURL, config);
    Logger.log("getResponseCode() "+response.getResponseCode());
    Logger.log("getContentText() "+response.getContentText().toString());
    jResponse = JSON.parse(response)
    Logger.log("jResponse: "+jResponse)
    Logger.log("refID"+jResponse.reference)

}

// string.replace('/\n/g', '<p>') isnt working for me? why?
function test_text_fixxer() {
  let description="line1\nline2\nline3\n"
  if(/\n/.test(description)) {
    Logger.log("found crlf")
  }
  Logger.log("description="+description)
  description.replace('/\n/g','<p>')
  Logger.log(" fixed="+description)
}
