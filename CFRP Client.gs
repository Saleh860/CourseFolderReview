/************************************************************************

                     Installation Procedure

Please check the generated log file, or if that failed, the Logger
output (Ctrl+Enter) for information about the status and progress of
installation procedure. 

To ensure successful installation make sure you perform the following
actions before running "install"
1- Make a copy of the Course Folder Review Process sheet in a new folder
2- Edit the script associated with the new sheet and place the new sheet
file Id in the configurationFileId below
3- Publish the script as a WebApp
For more detailed installation instructions please check the associated
documentation.
************************************************************************/

var configurationFileId = 'Place spreadsheet file Id here'; 

function install() {
  Install(configurationFileId);
  return;
}

/****************************************************/
/*  Objects and Functions imported from 
    Course Folder Review Library                    */
/****************************************************/

var Config = CFRLib.Config;
var Log = CFRLib.Log;
var Dataset = CFRLib.Dataset;
var ReviewProcess = CFRLib.ReviewProcess;
var SummaryTable = CFRLib.SummaryTable;
var Submissions = CFRLib.Submissions;
var Install = CFRLib.Install;
var WebCallbacks = CFRLib.WebCallbacks;

/******************************************************
periodicTrigger should be called every few hour for
sending review action emails (comments/approval)
Also moves approved files to the proper course folders
******************************************************/
function periodicTrigger() {
  //Load custom configuration for this client
  Config.load(configurationFileId);
  
  //Process reviewed items:
  //1. Send Comments --> Awaiting Response
  //2. Approved --> Complete
  var reviewProcess = ReviewProcess();
  
  //Scan the entire ReviewSheet for potential actions
  var actions = reviewProcess.getActions();
  
  //Peform the corresponding actions
  actions.forEach(function(action){ action.perform(); });
  
  //Send queued emails
  reviewProcess.finalize();
}

/******************************************************
formSubmitTrigger should be called for each new form 
submission to properly rename course file sections and
move uploaded files to the review folder and replace 
earlier versions of the uploaded files, if any. 
******************************************************/
function formSubmitTrigger() {
  //Load custom configuration for this client
  Config.load(configurationFileId);
  
  //Obtain new submissions
  var submissions = Submissions();
  
  //Insert new submitted files into the review process
  submissions.processAll();
}

/******************************************************
doGet is the Web App entry point. It renders a table 
that summarizes the status of each course folder 
section and provides links to each corresponding file 
******************************************************/
function doGet() {
  //Load custom configuration for this client
  Config.load(configurationFileId);
  
  //Convert the ReviewSheet to HTML presentation
  return SummaryTable.doGet();
}

/******************************************************
sendCommentsCallback is called when the user clicks on 
the "No" button to reject a submitted course section
and send comments to the user who submitted it.
******************************************************/
function sendCommentsCallback(semester, course, section, cellId) {
  //Load custom configuration for this client
  Config.load(configurationFileId);
  
  //Invoke the "sendComments" callback function to mark
  //the corresponding ReviewSheet cell as "Send Comments"
  return WebCallbacks.sendComments(semester, course, section, cellId);
}

/******************************************************
approveCallback is called when the user clicks on 
the "OK" button to approve a submitted course section 
******************************************************/
function approveCallback(semester, course, section, cellId) {
  //Load custom configuration for this client
  Config.load(configurationFileId);

  //Invoke the "approve" callback function to mark
  //the corresponding ReviewSheet cell as "Approved"
  return WebCallbacks.approve(semester, course, section, cellId);
}
