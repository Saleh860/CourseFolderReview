/********************************************************************/
/*                     Functor Helpers                              */
/*        Imported from Automated Workflow Library                  */
/********************************************************************/

/********************************************************************
* Partial application helper
*
* Returns a function with all parameters fixed except the last one. 
* The fix values are given as an array x
* 
********************************************************************/
var _apply = AW._apply;

var  _eq = AW._eq;

var  _lt = AW._lt;

var  _le = AW._le;

var  _ne = AW._ne;

var  _id = AW._id;

var  _property = AW._property;

var  _at = AW._at;

var  _or = AW._or;

var  _and = AW._and;

var  _compose = AW._compose;

var  _between = AW._between

var  _proxy = AW._proxy;

/********************************************************************/
/*                         Test Helpers                             */
/*        Imported from Automated Workflow Library                  */
/********************************************************************/

/********************************************************************
* Evaluate the given test case and if it returns false, print out 
* the code of the failed test case.
* A successful test case function must return true. 
********************************************************************/
var test = AW.test;

/********************************************************************/
/*                         Log Helpers                              */
/*        Imported from Automated Workflow Library                  */
/********************************************************************/

var Log = AW.Log;


/********************************************************************/
/*                      File/Folder Helpers                         */
/*        Imported from Automated Workflow Library                  */
/********************************************************************/

/*********************************************************************
* File object container with delayed evaluation
* 
* has two members:
* getId() returns the Id
* getObj() returns the DriveApp object
*********************************************************************/
var File = AW.File; 

/*********************************************************************
* Folder object container with delayed evaluation
* 
* has two members:
* getId() returns the Id
* getObj() returns the DriveApp object
*********************************************************************/
var Folder = AW.Folder; 

/********************************************************************/
/*                       Mail Helpers                               */
/*        Imported from Automated Workflow Library                  */
/********************************************************************/

var MailQueues = AW.MailQueues;

/********************************************************************/
/*                         Dataset                                  */
/*        Imported from Automated Workflow Library                  */
/********************************************************************/

var Dataset = AW.Dataset;

/********************************************************************/
/*                         Serialize                                */
/*        Imported from Automated Workflow Library                  */
/********************************************************************/
var serialize = AW.serialize;

/********************************************************************/
/*                         Settings                                 */
/*        Imported from Automated Workflow Library                  */
/********************************************************************/
var Settings = AW.Settings; 

/********************************************************************/
/********************************************************************/

/********************************************************************
Submissions handles recently submitted course files
********************************************************************/

function Submissions() {
  var processRecord;
  var loadRecords;
  var addRequestForReview;
  var mailQueues = MailQueues('Course Folders Multiple Notifications');

  /********************************************************************
  *                          Public Interface
  ********************************************************************/

  var submissions = {
    /********************************************************************
    * Scan submitted course files, adjust their names,
    * place them in the review folder and add a review record 
    * marking them as (waiting for review)
    ********************************************************************/
    processAll: function() {
      
      Log.info(function(){return ['Entering Function',"Sumbissions.processAll"]});
      
      var records = loadRecords(Config.submissionFormId);
      
      if(records) {
        /*Avoid writing to files while other processes are running*/
        var concurrency = Concurrency();
        if(concurrency.waitForTurn()) {
          Log.info(function(){return ['Submissions.processAll',"Found " + records.length + " new submissions."]});
          
          records.map(function (record) {
            processRecord(record);
            
            if(record.timestamp > Config.getLastTimestamp())
            Config.setLastTimestamp(record.timestamp);
          });
          
          mailQueues.sendQueuedEmails();
        }
        
      } else {
        Log.error(function(){
          return ['Submissions.processAll',"Invalid submission form, id=" + 
                  Config.subissionFormId]});
      }
      
      Log.info(function(){return ['Leaving Function',"Sumbissions.processAll"]});
    },
    
    testloadRecords: function() {
      Log.info(function(){return ['Test','Testing loadSubmissionRecords']});
      Config.load('1VITxYdTsBnll-7a9QWvwg0H30TI4lIpEfthNfmyiZPk');
      var records =   loadRecords(Config.submissionFormId);
      Log.info(function(){return ['Dump', records]});  
    },
  };

  /*****************************************************************
  *                      Private Members                           *
  *****************************************************************/
  
  /********************************************************************
  * load list of submitted files from Form with given fileId
  ********************************************************************/
  loadRecords = function(fileId) {
    Log.info(function(){return  ['Entering Function',"Submissions.loadRecords"]});
    var records = null;
    
    //Open spreadsheet
    try {
      var form = FormApp.openById(fileId);    
      var lastTimestamp = Config.getLastTimestamp();
      lastTimestamp.setTime(lastTimestamp.getTime()+1);
      var selectSemesterItem = form.getItemById(Config.selectSemesterItemId);
      var selectCourseItems = Config.selectCourseItemIds.map(function(id){return form.getItemById(id);});
      var fileUploadItems = Config.fileUploadItemIds.map(function(id){return form.getItemById(id);});
      
      records = form.getResponses(lastTimestamp).map(
        function(response){
          var semesterName = response.getResponseForItem(selectSemesterItem).getResponse();
          var semesterIndex = -1;
          Config.semesters.forEach(function(semester,i){
            if(semester.name == semesterName)
              semesterIndex = i;
          });
          
          var course = parseCourseFullName(
            response.getResponseForItem(
              selectCourseItems[semesterIndex]).getResponse());
          
          var file = null;
          var fileIds = fileUploadItems.map(function(item) {
            var r = response.getResponseForItem(item);

            if(r) {
              var fileId = r.getResponse()[0];//Only one file allowed per section
              if(fileId){
                try{
                  file = Drive.Files.get(fileId);
                } catch(e) {
                  Log.warn(function(){return ['Submissions.loadRecords','Invalid FileId='+fileId]});
                  fileId = null;
                }
              }
              return fileId;
            }
            return null;
          });
          
          return {
            timestamp: response.getTimestamp(), 
            email: (response.getRespondentEmail()? response.getRespondentEmail() : 
                    (file && file.lastModifyingUser?file.lastModifyingUser.emailAddress:'')),
            semester: Config.semesters[semesterIndex],
            course: course, 
            fileIds: fileIds,
          };
        });
      
      
    } catch(e) {
      Log.error(_proxy(['Submissions.loadRecords','Failed to load submission database\n' + e.toString()]));
    }
    Log.info(function(){return  ['Leaving Function',"Submissions.loadRecords"]});
    return records;
  };

  /********************************************************************
  * Rename submitted course file and place it in the review folder
  ********************************************************************/
  processRecord = function(record) 
  {
    Log.info(function(){return ['Entering Function',"Submissions.processRecord" + serialize(record)]});
    
    record.fileIds.forEach(function(fileId, i){
      if(fileId) {
        record.section = Config.sections[i];
        record.fileId = fileId;
        //1. Rename to standard filename
        File(fileId).rename(
          standardCourseFileName(
            {semesterCode: record.semester.code,
             courseCode: record.course.code,
             sectionFullName:(record.section.code.toString()+
                              '-'+record.section.name.toString()),
             courseName: record.course.name}));
        
        /* Add review request to the Review Sheet */
        addRequestForReview(record);
      }
    });
    
    Log.info(function(){return  ['Leaving Function',"Submissions.processRecord"]});
  };

  /********************************************************************
  * Insert a review request record for the newly processed 
  * submitted file 
  ********************************************************************/
  addRequestForReview = function (record) {
    Log.info(function(){return ["Entering Function" , "Submissions.addRequestForReview("+ 
                               [record.semester.code,record.course.code,record.section.code].join("_") + ")"]}); 
    
    var fileReplacementMessage = 'The new file will replace an existing file. '+
                    'The old file will be moved to the "Review/Trash" folder.';
    
    var sheet = SpreadsheetApp.openById(Config.reviewSheetFileId)
    .getSheetByName(record.semester.name); 
    
    var dataset = AW.Dataset(sheet, 'Code');
    var courseRecord = dataset.selectRecords(
      function(r){return r.getValue('Code')==record.course.code});
    if(courseRecord.length) {
      try {
        //TODO: Warn if replacing a file and old file status "Approved" or "Completed"
        var oldFileHyperlink = courseRecord[0].getValue(record.section.code);
        var oldFile = File(parseHyperlink(oldFileHyperLink).url);
        //Existing course file  
        //Move the old file in the existing course row to trash
        if(oldFile.getId()!=record.fileId) {
          Log.info(function(){
            return ['Submissions.addRequestForReview', 
                    fileReplacementMessage] });
          oldFile.moveToFolder(Folder(Config.trashFolderId));
        }
      } catch(e) {
        //It's okay if no old course file exists or if it cannot be parsed correctly 
      }
      //Put course file in Review folder
      File(record.fileId).moveToFolder(Folder(Config.reviewFolderId));
      
      //Update course row
      courseRecord[0].setValue(record.section.code,  
                               makeHyperlink('https://drive.google.com/open?id=' + 
                                             record.fileId, record.timestamp));
      courseRecord[0].setValue(record.section.name, 'Review');
      courseRecord[0].setValue('Coordinator',record.email); 
      
      var subject = "File [" + record.courseName +  "] " + record.section ;
      var message = "Thank you for submitting the course file. \r\n" +
        "The file has been received and is awaiting review.\r\n\r\n" +
          fileReplacementMessage +
            "The new file is located at " + record.fileUrl;
      /*  mailQueues.enqueueEmail(record.email, subject, message,
      {name: 'Automatic Emailer'}); */
    }
    else { 
      Log.error(_proxy(['Submissions.addRequestForReview', 'Course Code ' + record.course.code + '  not found in Review Sheet of semester ' + record.semester.name]));
    }
    
    Log.info(function(){return ["Leaving Function", "Submissions.addRequestForReview"]}); 
  };
  
  return submissions;  
}

/********************************************************************/
/********************************************************************/

/********************************************************************
ReviewProcess encapsulates the review process workflow
********************************************************************/
function ReviewProcess() {
  var SendCommentsAction;
  var ApprovedAction;
  var mailQueues = MailQueues('Course Folders Multiple Notifications');
  
  Config.loadCoursesUrls();
  
  var reviewProcess = {
    /********************************************************************
    * Find all files marked as 'Send Comments' or 'Approved' 
    ********************************************************************/
    getActions: function () {
      Log.info(function(){return ["Entering Function","ReviewProcess.getActions"]});
      
      var actions = new Array();
      
      var book = SpreadsheetApp.openById(Config.reviewSheetFileId);
      book.getSheets().forEach(function(sheet){
        Config.semesters.forEach(function(semester){
          if(semester.name==sheet.getName()) {
    
            var dataset = AW.Dataset(sheet, 'Code');
            var records = dataset.selectRecords(function(record) {return true;});
      
            records.forEach(
              function(record){
                Config.sections.forEach(
                  function(section){
                    switch(record.getValue(section.name)){
                      case 'Send Comments':
                        actions.push(SendCommentsAction(semester, record, section));
                        break;
                      case 'Approved':
                        actions.push(ApprovedAction(semester, record, section));
                        break;
                      default:
                        break;
                    }
                  });
              });
          }
        });
      });
      
      Log.info(function(){return ["Leaving Function", "ReviewProcess.getActions"]});
      
      return actions;
    },
    finalize: function() {
      mailQueues.sendQueuedEmails();
    }
  };

  var initializeAction = function(semester, record, section) {
    var action = {
      semesterCode: semester.code,
      sectionCode: section.code,
      courseCode: record.getValue('Code'),
      coordinator: record.getValue('Coordinator'),
      sectionFullName: (section.code+'-'+section.name),
      courseFullName: (makeCourseFullName(record.getValue('Code'),record.getValue('Name'))),
      fileId: (parseHyperlink(record.getFormula(section.code)).url.split('=')[1]),
      setStatus: function(newStatus) {record.setValue(section.name, newStatus);},
    }
    Log.info(function(){return ["ReviewProcess.initializeAction" ,AW.serialize(action)]});
    return action;
  }
  
  /*******************************************************************
  * A file marked as Approved should be moved to the Course Folder
  * And the lucky owner should be notified of the achievement
  *******************************************************************/
  ApprovedAction = function(semester,record,section) {
    var action = initializeAction(semester, record, section);
        
    action.perform = function() {
      Log.info(function(){return ["Entering Function" ,'ApprovedAction.perform']}); 

      var targetFolder = Folder(Config.courses.getByCode(action.courseCode).folderId);
      
      if(!targetFolder.exists()) {
        targetFolder = Folder(Config.coursesFolderId)
        .openCreatePath([action.courseFullName]);
      }
      
      if(!targetFolder) {
        Log.error(function(){return ['ApprovedAction', "Can't open or create the Course Folder for " + 
                                    action.courseFullName]; });
        }
      else {
        var file = File(action.fileId);
        if(!file.exists()) {
          Log.error(function(){return ['ApprovedAction', "Can't open Course file with id= " + 
                                      action.fileId]; });
        }
        else {
          file.moveToFolder(targetFolder);
            
          Log.info(function(){return "Approved file moved to course folder: " + 
            targetFolder.getName() });
            
          var subject = "Approved Section " + action.sectionCode + " of " +  action.courseCode;
          var message = "Thank you for submitting Section "+ 
            action.sectionFullName +" of the " + action.courseFullName + "course file. \r\n"+
              "The file has been reviewed and approved.\r\n\r\n"+
                "The file is now located in the following Course Folder: " + 
                  targetFolder.getUrl();
            
          mailQueues.enqueueEmail(action.coordinator, subject, message,
                                  {name: 'Automatic Emailer'}); 
            
          action.setStatus('Completed');
        }
        Log.info(function(){return ["Leaving Function", "ApprovedAction.perform"]}); 
      }
    };
    return action;
  }
  
  /*******************************************************************
  * A file is marked as "Send Comments" to enduce the workflow 
  * to notify the owner of the file to respond to the review comments
  * Once such a file is processed and the notification is sent,
  * the file is automatically marked as "Awaiting Response"
  *******************************************************************/
  SendCommentsAction = function(semester,record,section) {
    var action = initializeAction(semester, record, section);
    
    action.perform = function() {
      Log.info(function(){return ["Entering Function" ,'SendCommentsAction.perform']}); 

      var file = File(action.fileId);
      if(!file.exists()) {
        Log.error(function(){return ['SendCommentsAction', "Can't open Course file with id= " + 
                                    action.fileId]; });
      }
      else {
        //Put course file in Review folder
        file.moveToFolder(Folder(Config.reviewFolderId));

        var subject = "Comments on Section " + action.sectionCode + " of " +  action.courseCode;
      
        var message = "Thank you for submitting Section "+ 
          action.sectionFullName +" of the " + action.courseFullName + "course file. \r\n"+
            "There are some comments on your file that should be addressed.\r\n"+
              "Please review the comments and respond to them and if necessary resubmit "+
                "the course file.\r\n\r\nThe file with comments is located at " + 
                  file.getUrl() + '\r\n---' + Config.submissionFormUrl ;
      
        mailQueues.enqueueEmail(action.coordinator, subject, message,
                                {name: 'Automatic Emailer'});
      
      
        action.setStatus('Awaiting Response');
        
      }
    
      Log.info(function(){return ["Leaving Function", "SendCommentsAction.perform"]}); 
    };
    return action;
  }
  
  return reviewProcess;
}

/********************************************************************/
/********************************************************************/

/********************************************************************
Config encapsulates configuration persistence
********************************************************************/
var Config = {
  status: null,
  load: function (configurationFileId) {
    Log.enable(); //Enable logging to script logger only.
    try {
      var settingsSheet = SpreadsheetApp.openById(
        configurationFileId).getSheetByName('Settings');
      if(settingsSheet)
        var settings = AW.Settings(settingsSheet);
      else
        throw 'No sheet named Settings can be found';
    }
    catch(e) {
      AW.Log.error(
        AW._proxy(
          ['Config.load',
           'Failed to open "Settings" sheet from file with id = ' +
           configurationFileId + '\n' + e]));
      return;
    }
    
    Config.configurationFileId = configurationFileId;

    Config.logFileId = settings.get('LogFileId');
    Log.enable(Config.logFileId); //Enable logging to the log file
    
    Config.semesters = eval(settings.get('Semesters'));
    Config.courses = eval(settings.get('Courses'));
    Config.sections = eval(settings.get('Sections'));
    Config.submissionFormId = settings.get('SubmissionFormId');
    Config.submissionFormUrl = settings.get('SubmissionFormUrl');
    Config.coursesFolderId = settings.get('CoursesFolderId');
    Config.trashFolderId = settings.get('TrashFolderId');
    Config.reviewSheetFileId = settings.get('ReviewSheetFileId');
    Config.reviewFolderId = settings.get('ReviewFolderId');
    Config.getSectionName = function(code) {
      var section = this.sections.filter(function(section){return section.code==code;});
      if(!section) return null;
      return section.name;
    };
    Config.webAppUrl = settings.get('webAppUrl');
    Config.set = settings.set;
    Config.get = settings.get;
    Config.getLastTimestamp = 
      function(){return settings.get('LastTimestamp');};
    Config.setLastTimestamp = 
      function(newTimestamp) { settings.set('LastTimestamp',newTimestamp);};
    
    Config.selectSemesterItemId = settings.get('SelectSemesterItemId');
    Config.selectCourseItemIds = eval(settings.get('SelectCourseItemIds'));
    Config.fileUploadItemIds = eval(settings.get('FileUploadItemIds'));
    
    var addGetByCode = function(array){
      array.getByCode = function(code) {
        var obj=null;
        for(var i in array){
          if(array[i].code==code) {
            obj=array[i];
            break;
          }
        }
        return obj;
      }
    }
    
    var addGetByName = function(array){
      array.getByName = function(name) {
        var obj=null;
        for(var i in array){
          if(array[i].name==name) {
            obj=array[i];
            break;
          }
        }
        return obj;
      }
    }
    
    addGetByName(Config.semesters);
    addGetByCode(Config.semesters);
    addGetByName(Config.sections);
    addGetByCode(Config.sections);
    addGetByName(Config.courses);
    addGetByCode(Config.courses);
    
    Config.loadCoursesUrls = function() {
      Config.courses = eval(settings.get('CoursesWithUrls'));
      addGetByCode(Config.courses);
      addGetByName(Config.courses);
      Config.loadCoursesUrl = function(){};
    }
    
    Config.status = 'ready';    
  }
}


/********************************************************************/
/********************************************************************/

/********************************************************************
* Compose course file name in standard form
* record must contain the following fields
* semesterCode, courseCode, sectionFullName, courseName
********************************************************************/
function standardCourseFileName(record) {

  return [record.semesterCode, record.courseCode, 
          record.sectionFullName, record.courseName].join('_');
}

/********************************************************************
* Parse course file name in standard form
* returns a record containing the following fields
* semesterCode, courseCode, sectionFullName, courseName
********************************************************************/
function splitStandardCourseFileName(filename) {
  a = filename.split('_');
  return {semesterCode:a[0], courseCode:a[1], sectionFullName:a[2], courseName:a[3]};
}

function makeHyperlink(url,text) {
  return '=HYPERLINK("'+url+'","'+text+'")';
}
function parseHyperlink(h) {
  return {
    url : h.split('HYPERLINK("')[1].split('","')[0],
    text : h.split('","')[1].split('")')[0]
  };
}

function makeCourseFullName(code, name) {
  return code+' ['+name+']';
}
function parseCourseFullName(fullName) {
  var parts = fullName.split(' [');
  return {code:parts[0],
          name:parts[1].slice(0,parts[1].length-1)};
}
function testCourseFullName() {
  Logger.log(makeCourseFullName('123456-3','Some Course Name'));
  Logger.log(AW.serialize(parseCourseFullName(makeCourseFullName('123456-3','Some Course Name'))));
  return;
}

function testParseHyperlink() {
  Logger.log(AW.serialize(parseHyperlink(makeHyperlink('https://docs.google.com/','00/00 00:00'))));
}

/********************************************************************/
/********************************************************************/

function Concurrency() {
  
  //Temporarily Disabled
  return {
    getTurn: function() {
      return 1;
    },
    
    wait: function(){
    },
    waitForTurn: function() {
      return true;
    },
    yield: function() {
    }
  };
  
  var sheet = SpreadsheetApp.openById(config().concurrencySheetId);
  var turnsSheet = sheet.getSheetByName('Turns');
  var signature = Math.round(Math.random()*1000000000000);
  var turn = turnsSheet.getLastRow()+1;
  var currentTime=new Date();
  var lastRow = turnsSheet.appendRow([currentTime, signature]).getLastRow();
  while(turn<=lastRow){
    if(turnsSheet.getSheetValues(turn, 2, 1, 1)[0][0]==signature)
      break;
    turn++;
  }
  var currentSheet = sheet.getSheetByName('CurrentTurn');
  var currentTurn = {
    get: function() {
      return currentSheet.getSheetValues(1, 1, 1, 1)[0][0]; 
    },
    set: function(t){
      Logger.log(currentSheet.getRange(1,1).getValue())
      currentSheet.getRange(1, 1).setValue(t);
    }
  };
  
  return {
    getTurn: function() {
      return turn;
    },
    
    wait: function(){
      Logger.log(currentTime.getTime());
      Logger.log(sheet.getSheetValues(
          currentTurn.get(), 1, 1, 1)[0][0].getTime());
      if(!currentTurn.get() || 
         currentTime.getTime() -
        sheet.getSheetValues(
          currentTurn.get(), 1, 1, 1)[0][0].getTime()>240000)
        currentTurn.set(turn);
      
      while(currentTurn.get()!=turn){
        Utilities.sleep(1000);
      }
    },
    waitForTurn: function() {
      return true;
    },
    yield: function() {
      currentTurn.set(turn+1);
    }
  }
}

function testConcurrency() {
  var concurrency = Concurrency();
  Logger.log(concurrency.getTurn());
  concurrency.wait();
  Utilities.sleep(5000);
  concurrency.yield();
}

/********************************************************************/
/********************************************************************/

/*
This module prepares a table for each active semester containing
a row for each course, with a column for each section indicating
the status of the section (missing, submitted, rejected, approved)
*/
var Summary_CoursesFolderId = '1N6FUXnfaUfM2XesRnGuNMC112OVXnLtj';
var Summary_ActiveSemesters = ['17F','18S'];
var Summary_Sections = [['0', 'Contents'],
                            ['A', 'Integrity'],
                            ['B', 'Day1'],
                            ['C', 'Syllabus'],
                            ['DG', 'Plan'],
                            ['H', 'Activities'],
                            ['I', 'Midterms'],
                            ['J', 'Final'],
                            ['K', 'Survey'],
                            ['L', 'CAR']];
var Summary_Courses= [
    [
        {code:'803213-3', name:'Introduction to Computer Programming',report: true}, 
        {code:'803223-1', name:'Professional Ethics',report: true}, 
        {code:'803302-3', name:'Electromagnetic Fields',report: true}, 
        {code:'803303-3', name:'Electronics (1)',report: true}, 
        {code:'803318-3', name:'Digital Logic Design',report: true}, 
        {code:'803411-3', name:'Electrical Power Systems (1)',report: true}, 
        {code:'803412-3', name:'Computer Architecture',report: true}, 
        {code:'803413-3', name:'Signals and Systems Analysis',report: true}, 
        {code:'803414-3', name:'Computer Networks',report: true}, 
        {code:'803415-3', name:'Electrical and Electronic Measurements',report: true}, 
        {code:'803514-3', name:'Renewable Energy Systems',report: true}, 
        {code:'803519-3', name:'Protection against Job Hazards',report: true}, 
        {code:'803533-3', name:'Electromechanical Conversion (2)',report: true}, 
        {code:'803551-3', name:'Communication Systems (2)',report: false}
    ], 
    [
        {code:'803224-3', name:'Basics of Electric Circuits',report: true}, 
        {code:'802226-3', name:'Introduction to Engineering Design',report: true}, 
        {code:'803324-3', name:'Electric Circuits',report: true}, 
        {code:'803326-3', name:'Electronics (2)',report: true}, 
        {code:'803322-3', name:'Electromechanical Energy Conversion (1)',report: true}, 
        {code:'803325-3', name:'Programming (2)',report: true}, 
        {code:'803423-3', name:'Control Engineering',report: true}, 
        {code:'803422-3', name:'Communication Systems (1)',report: true}, 
        {code:'803425-3', name:'Power Electronics',report: true}, 
        {code:'803424-3', name:'Electical Installations',report: true}, 
        {code:'803421-3', name:'Optical Engineering',report: true}, 
        {code:'803522-3', name:'Neural Networks and Fuzzy Systems',report: true}, 
        {code:'803521-2', name:'Engineering Management',report: true}, 
        {code:'803532-3', name:'Protection',report: true}, 
        {code:'803530-3', name:'High Voltage',report: true}, 
        {code:'803428-3', name:'Electronic Circuits',report: true}, 
        {code:'803573-3', name:'Logic Controllers and Microcontrollers',report: true}, 
        {code:'803534-3', name:'Special Machines',report: false}, 
        {code:'803523-3', name:'Biomedical Engineering',report: true}, 
        {code:'803427-3', name:'Computer Programming (3)',report: true}, 
        {code:'803429-3', name:'Power Systems (2)',report: false}
    ]
];

function Summary_Status(id) {
  switch(id) {
    case 0:
      return 'missing';
    case 1:
      return 'submitted';
    case 2:
      return 'rejected';
    case 3:
      return 'approved';
    default:
      return 'unknown';
  }
}

function Summary_Section(code, name, fileUrl, status) {
  if(typeof(status)=='number') status = Summary_Status(status);
  return {url:fileUrl, 
          status: status,
          sheetRow: -1,
          code: code,
          name: name,
/*           isReady:status=='ready',
          isSubmitted: status=='submitted',
          isRejected: status=='rejected',
          isApproved:status=='approved' */
         };  
}

function Summary_Course(courseCode, courseName, folderUrl,sections) {
  return {code: courseCode, 
          name: courseName,
          url: folderUrl, 
          sections: (sections?sections:new Array())};
}

function Summary_Semester(semester,sections,courses) {
  return {
    semester: semester,
    sections: sections,
    courses:courses
  };
}

var Summary_TableCached = null;

// via: http://blog.stchur.com/2007/04/06/serializing-objects-in-javascript/
function serialize(_obj)
{
   // Let Gecko browsers do this the easy way
   if (typeof _obj.toSource !== 'undefined' && typeof _obj.callee === 'undefined')
   {
      return _obj.toSource();
   }
   // Other browsers must do it the hard way
   switch (typeof _obj)
   {
      // numbers, booleans, and functions are trivial:
      // just return the object itself since its default .toString()
      // gives us exactly what we want
      case 'number':
      case 'boolean':
      case 'function':
         return _obj;
         break;

      // for JSON format, strings need to be wrapped in quotes
      case 'string':
         return '\'' + _obj + '\'';
         break;

      case 'object':
         var str;
         if (_obj.constructor === Array || typeof _obj.callee !== 'undefined')
         {
            str = '\n[';
            var i, len = _obj.length;
            for (i = 0; i < len-1; i++) { str += serialize(_obj[i]) + ','; }
            str += serialize(_obj[i]) + ']';
         }
         else
         {
            str = '\n{';
            var key;
            for (key in _obj) { str += key + ':' + serialize(_obj[key]) + ','; }
            str = str.replace(/\,$/, '') + '}';
         }
         return str;
         break;

      default:
         return 'UNKNOWN';
         break;
   }
}
function Summary_MakeEmptyTable() {
  var reviewFolder = DriveApp.getFolderById(Summary_CoursesFolderId);
  var table = Summary_Courses.map(function(semester,i){
    return Summary_Semester(
      Summary_ActiveSemesters[i],
      Summary_Sections.map(function(section){
        return {code: section[0],name:section[1]};            
      }),
      semester.map(function(course){
        var it = reviewFolder.getFoldersByName(course);
        var folderId=
            it.hasNext()?
            reviewFolder.getFoldersByName(course).next().getUrl():null;
        return Summary_Course(
          course.code, 
          course.name,
          folderId, 
          Summary_Sections.map(function(section){
            var status=(section[1]!='CAR')?0:(course.report?0:4);
            return Summary_Section(section[0], section[1], '', status);
          }));
      })); 
  });
  return table;
}

function Summary_All() {
  var database = loadSubmissionDb(config().submissionSpreadsheetId);
  var cachedTable = database.settings.get('EmptySummaryTable');
  var table;
  if(cachedTable)
    table = eval(cachedTable);
  else {
    table = Summary_MakeEmptyTable();
    database.settings.set('EmptySummaryTable',serialize(table));
  }
  var semesterId = function(semester) {
    return Summary_ActiveSemesters.indexOf(semester);
  }
  
  var courseCodes = Summary_Courses.map(function(sem) {
    return sem.map(function(course){
      return course.code;
    });
  });
  
  var courseId = function(semester,courseCode) {
    var id = courseCodes[semesterId(semester)].indexOf(courseCode);
    return id;
  }
  var sectionCodes = Summary_Sections.map(function(section){
    return section.join('-');
  });
  var sectionId = function(sectionCode) {
    return sectionCodes.indexOf(sectionCode);
  }
  
  table.getCourseSection = function(semester,courseCode,sectionCode) {
    return table[semesterId(semester)]
    .courses[courseId(semester,courseCode)]
    .sections[sectionId(sectionCode)]
  }
  
  var records = database.getAllRecords();
  records.forEach(function(record) {
    if(record.fileUrl) {
      var section = table.getCourseSection(record.semester,record.courseCode,record.section);
      section.timestamp = record.timestamp;
      section.url = record.fileUrl;
      section.status = 'submitted';
    }
  });
  
  var reviewSheets = getReviewSheets();
  reviewSheets.forEach(function(sheet) {
    var reviewRecords = getReviewRecords( sheet.file );
    reviewRecords.forEach(
      function (record) {
/*        Logger.log(sheet);
        Logger.log(record); */
        var section = table.getCourseSection(sheet.semester,record.courseCode,sheet.section);
        section.url = record.fileUrl;
        section.sheetRow = record.rowIndex;
        switch(record.status) {
          case 'Send Comments':
          case 'Awaiting Response':
            section.status = 'rejected';
            break;
          case 'Completed':
          case 'Approved':
            section.status = 'approved';
            break;
          case 'Review':
            section.status = 'submitted';
            break;
          default:
            section.status = 'unknown';
            break;
        }
    });
  });
  
  Logger.log(table);
  return table;
}

function Summmary_getReviewRecord(semester, section, rowId) {
  var reviewFile = DriveApp.getFolderById(config().reviewFolderId)
  .getFoldersByName(semester).next()
  .getFoldersByName(section).next()
  .getFilesByName(config().reviewChecklistFilename).next();
  var records = getReviewRecords( File(reviewFile) , rowId, 1);
  return records[0];
}

function Summary_sendComments(semester, section, rowId, cellId) {
  var msg = 'Send Comments:'+semester+','+section+','+rowId;
  Log.info(function(){return ['SummaryTableAction',msg]});
  var record = Summmary_getReviewRecord(semester, section, rowId);
  record.updateStatus('Send Comments');
  return cellId;
}

function Summary_approve(semester, section, rowId, cellId) {
  var msg = 'Approve:'+semester+','+section+','+rowId;
  Log.info(function(){return ['SummaryTableAction',msg]});
  var record = Summmary_getReviewRecord(semester, section, rowId);
  record.updateStatus('Approved');
  return cellId;
}


/********************************************************************/
/********************************************************************/

var SummaryTable= {
  doGet: function() {
    Log.info(function(){return ['SummmaryTable','Requested']})
    
    Config.loadCoursesUrls();
    
    var t = HtmlService.createTemplateFromFile('SummaryTable');
    
    t.editable = Drive.Files.get(Config.reviewSheetFileId).editable;
    
    t.sheets = SpreadsheetApp.openById(Config.reviewSheetFileId).getSheets();
    t.statusClass = function(status) {
      switch(status) {
        case 'Send Comments':
        case 'Awaiting Response':
          return 'rejected';
        case 'Completed':
        case 'Approved':
          return 'approved';
        case 'Review':
          return 'submitted';
        case 'Missing':
          return 'missing';
        default:
          return 'unknown';
      }
    };
    
    t.formatTime = function(ts) {
      return ts.getDate() + '/' + (ts.getMonth()+1) + ' ' + 
        ts.getHours() + ':' + 
          (ts.getMinutes()<10?'0':'')+ts.getMinutes();
    };
    
    var h = t.evaluate();
    return h;
  }
};

function WebCallbacks() {  
  var getReviewRecord = function (semester, course) {  
    var reviewSheet = SpreadsheetApp.openById(Config.reviewSheetFileId).getSheetByName(semester);
    var dataset = Dataset(reviewSheet, 'Code');
    var records = dataset.selectRecords(function(r){return r.getValue('Code')==course;});
    return records[0];
  }
  
  return {
    sendComments: function(semester, course, section, cellId) {
      var msg = 'Send Comments:'+semester+','+course+','+section;
      Log.info(function(){return ['WebCallbacks',msg]});
      var record = getReviewRecord(semester, course);
      record.setValue(section,'Send Comments');
      return cellId;
    },
    approve: function(semester, course, section, cellId) {
      var msg = 'Approve:'+semester+','+course+','+section;;
      Log.info(function(){return ['WebCallbacks',msg]});
      var record = getReviewRecord(semester, course);
      record.setValue(section,'Approved');
      return cellId;
    }
  };
}

/********************************************************************/
/********************************************************************/

/* This is an installation script for the Course Folder Review Library
The script takes a configuration spreadsheet containing:
1. Active Semesters
2. Courses
3. Sections

The script creates and configures the following
1. Settings sheet
2. Log file
3. Submission Form
4. Review and Showcase Website
5. Periodic trigger to organizeCourseFolders

*/
function Install(configurationFileId) {
    
  //perform installation
  main();
  
  function main() {
    verifyClientScript();
    var installationFolder = getInstallationFolder();
    var cfgBook = getConfigurationBook();
    var settings = getSettings(cfgBook);
    enableLogFile(settings, installationFolder);
    var data = loadDataArrays(settings, cfgBook);
    createFolders(settings, installationFolder, data);
    var form = openCreateSubmissionForm(settings,installationFolder, data);
    createReviewSheet(settings,installationFolder, data);
    attachTriggers(form);
    deployWebApp(settings);
  }
  
    
  /****************************************************************/
  //Stop the execution of the script, displaying the given error message
  /****************************************************************/
  function error(message){
    Log.error(_proxy(['Install',message]));
    throw message;
  }
  
  /****************************************************************/
  //Make sure the script is being invoked from a client script
  //If the script is being run directly, throw an error.
  /****************************************************************/
  function verifyClientScript() {
    if('1Mt2WkUVxYOm9X8UozUcA-oww2EYY3HPT_nHwcfdR0ZdwdlQYw_ueTJDp'==ScriptApp.getScriptId()) 
      error('Course Folder Review Library Install ' +
      'script can only be executed from another Client script. Please create your '+
        'own copy of the Course Folder Review Client script and use it to call Install');
  }
  
  /****************************************************************/
  //Make sure the configuration file is properly located in 
  //a valid installation folder (non-root, owned by this user)
  /****************************************************************/
  function getInstallationFolder() {
    //Open configuration file 
    try {
      var cfgFile = DriveApp.getFileById(configurationFileId);
    }
    catch(e) {
      error('Can\'t open configuration file with id = ' +
            configurationFileId + '\n' + e);
    }

    var it = cfgFile.getParents();
    while(it.hasNext()) {
      var installationFolder = Folder(it.next());
      if(installationFolder.getId() != DriveApp.getRootFolder().getId()) {
        break;
      }
    }
    if(installationFolder.getId() == DriveApp.getRootFolder().getId()) {
      error('The Course Folder Review Process can\'t be set '+
            'up in the root folder of your Google Drive. Please create a subfolder '+
            'and move the configuration file there.');
    }
  
    if(installationFolder.getOwnerEmail() != Session.getEffectiveUser().getEmail()) {
      error('Installation folder must be owned by the '+
            'same user running the installation script. Place the installation '+
            'script in your own folder and try again');
    }
    Log.info(_proxy(['Install', 'Attempting to install the '+
                     'Course Folder Review Process by user ' + 
                     Session.getEffectiveUser().getEmail() + 
                     ' to folder: ' + installationFolder.getName()]));
    return installationFolder;
  }
  
  /****************************************************************/
  //Open configuration spreadsheet
  /****************************************************************/
  function getConfigurationBook() {
    try {
      var cfgBook = SpreadsheetApp.openById(configurationFileId);
    } 
    catch (e) {
      error('Configuration file must be a google Spreadsheet. '+
            'Please see the sample configuration sheet as an example.');
    }
    return cfgBook;
  }
  
  /****************************************************************/
  //Create or open Settings sheet in the configuration Book
  /****************************************************************/
  function getSettings(cfgBook){  
    try {
      var settingsSheet = cfgBook.getSheetByName('Settings');
      if(!settingsSheet) {
        Log.info(function(){return ['Install','Initializing a fresh installation'];});
        settingsSheet = cfgBook.insertSheet(cfgBook.getSheets().length);
        settingsSheet.setName('Settings');
        settingsSheet.appendRow(['Key','Value']);
      }
      else {
        Log.info(function(){return ['Install','Resuming a prior installation'];});
      }      
    }
    catch (e) {
      error('Install',
            'Failed to open or create Settings '+
            'sheet inside the Configuration Book.');
    }
    
    //Create Settings object
    var settings = Settings(settingsSheet);
    if(!settings) {
      error('Failed to load settings from sheet.');
    }
    return settings;
  }
  
  /****************************************************************/
  //Open or Create Log File
  /****************************************************************/
  function enableLogFile(settings, installationFolder) {
    //Try to Open log file
    var logFileId = settings.get('LogFileId');
    if(File(logFileId).exists()) {
      Log.enable(logFileId);
      Log.info(_proxy(['Install', 'An existing logFileId was found and will be used.']));
    }
    else {
      //Try to Create log file
      try {
        var logSpreadsheet = 
            SpreadsheetApp.create('Course Folder Review Process Log');
        logSpreadsheet.appendRow(
          ['Timestamp','Level','Type','Message']);

        logFileId = logSpreadsheet.getId();
      
        Log.enable(logFileId);
        Log.info(_proxy(['Install', 'A new Log File was created.']));

        File(logFileId).moveToFolder(installationFolder);
      
        settings.set('LogFileId',logFileId);
      }
      catch(e) {
        Log.warn(_proxy(['Install', 
                         'Failed to create log spreadsheet file.\n' + e]));
      }
    }
  }
    
  /****************************************************************/
  //Load data arrays (semesters, courses, sections) 
  //from configuration file
  /****************************************************************/
  function loadDataArrays(settings, cfgBook) {
    //In case of a resumed installation, this flag to keep track if any 
    //of the loaded data arrays is different from the last run of install 
    var data = {changed: false};

    //Load semesters array
    data.semesters = loadArray('Semesters',function(record){
      return {code: record.getValue('Code'), name: record.getValue('Name')};
    });
    
    //Load courses array
    data.courses = loadArray('Courses', function(record){
      var code = record.getValue('Code');
      var name = record.getValue('Name');
      return {code: code, 
              name: name,
              fullName: makeCourseFullName(code,name),
              semester: record.getValue('Semester'),
              coordinator: record.getValue('Coordinator')};
    });
    
    //Load sections array
    data.sections = loadArray('Sections', function(record){
      return {code: record.getValue('Code'), 
              name: record.getValue('Name')};
    });
    return data;

    /****************************************************************/
    //Load array from configuration file
    /****************************************************************/
    function loadArray(arrayName,mapRecordToObject) {
      try {
        var sheet = cfgBook.getSheetByName(arrayName);
        
        if(!sheet) {
          error('Bad configuration file format. Must contain "'+arrayName+'" sheet');
        }   
        var ds = Dataset(sheet,'Code');
        
        var array = ds.selectRecords(_proxy(1)).map(mapRecordToObject);
        
        var arraySetting = serialize(array);
        if(settings.get(arrayName)==arraySetting)
          Log.info(function(){return['Install',arrayName +' unchanged']});
        else {
          settings.set(arrayName, arraySetting);
          data.changed = true;
        }
      } 
      catch(e) {
        error('Bad configuration file format. Can\'t load "Semesters" sheet.\n'+e);
      }
      return array;
    }
  }
  
  function createFolders(settings, installationFolder, data) {
    //Create Approved Courses Folder
    var coursesFolder = bindFolderId(settings, 'CoursesFolderId', installationFolder, 'Courses');
    
    //Create Approved Course Folders
    data.courses.forEach(function(course){
      var courseFolder=coursesFolder.openCreateFolder(course.fullName);
      course.folderId = courseFolder.getId();
    });
    
    //Create Review Folder
    var reviewFolder = bindFolderId(settings, 'ReviewFolderId', installationFolder, 'Review');  
    
    //Create TrashFiles Folder
    var trashFolder = bindFolderId(settings, 'TrashFolderId', installationFolder, 'Trash Files');
  }

  /**************************************************************
  If the given folderId is valid, returns a corresponding 
  Folder object. Otherwise, returns null.
  **************************************************************/
  function getFolder(folderId) {
    var folder = Folder(folderId);
    if(folder.exists()) {
      return folder; 
    }
    return null;
  }
  
  /**************************************************************
  Internal function to check if an installation subfolder is 
  already set up and if not, set it up. 
  Returns a Folder object of the existing or new subfolder
  **************************************************************/
  function bindFolderId(settings, settingName,inFolder,folderName){
    var folderId = settings.get(settingName);
    var folder=getFolder(folderId);
    if(folder) {
      Log.info(function(){return ['Install','An existing "'+settingName+'" already bound.'];});
    }
    else {
      Log.info(function(){return ['Install','Binding "'+settingName+
                                  '" subfolder with name "'+folderName+'".'];});
      folder = inFolder.openCreateFolder(folderName);
      settings.set(settingName, folder.getId());
    }
    
    return folder;
  }
  
  /****************************************************************/
  //Open Submission Form if a valid one exists
  //Otherwise create a new Submission Form
  /****************************************************************/
  function openCreateSubmissionForm(settings, installationFolder, data){
    var form;
    if(!data.changed) {
      form = openSubmissionForm();
    }
    if(form) {
      Log.info(function(){return ['Install','An existing submissionFormId was found and will be used.'];});
    }
    else {
      form = createSubmissionForm();
    }
    return form;
    
    /****************************************************************/
    //Create Submission Form
    /****************************************************************/
    function openSubmissionForm() {
      var submissionFormId = settings.get('SubmissionFormId');
      if(File(submissionFormId).exists()) {
        try{
          return FormApp.openById(submissionFormId);
        }
        catch(e) {
          Log.warn(_proxy(['Install', 'Existing submissionFormId was invalid. A new form will be created.']));
          return null;
        }
      }
      return null;
    }
    
    function createSubmissionForm() {
      //Create from Template Form
      //var form = FormApp.create('Course Folder Submission Form');
      try {
        var formFile = DriveApp.getFileById('1TwbCrKuP3Om25ioNQI1JeW2OjpbVPxmZN3vGrWl2K60')
        .makeCopy('Course Folder Submission Form', installationFolder.getObj());
        
        var form = FormApp.openById(formFile.getId());
        
      }
      catch(e) {
        error('Cannot create submission form\n'+e);
      }
      
      var fileItem = form.getItems()[0];
      
      form.setDescription('Use this form to upload your course folder contents for review and approval by the Accreditation Committee.');
      var semestersItem = form.addMultipleChoiceItem().setTitle('Select Semester');
      var pages = new Array();
      var semestersChoices = new Array();
      var selectCourseItems = new Array();
      var fileUploadItems = new Array();
      data.semesters.forEach(function(semester) {
        var page = form.addPageBreakItem().setTitle(semester.name + ' Courses');
        selectCourseItems.push(
          form.addListItem().setTitle('Select Course')
          .setChoiceValues(
            data.courses.filter(function(course){
              return course.semester==semester.code;
            }).map(function(course){
              return course.fullName;
            })));
        
        semestersChoices.push(semestersItem.createChoice(semester.name, page));
        pages.push(page);
      });
      var filePage = form.addPageBreakItem().setTitle('Upload Files')
      .setHelpText('Only a single PDF or DOC file no larger than '+
                   '10MB is allowed per section. To merge multiple PDF files or '+
                   'compress them, you may use a compression service such as '+
                   'https://www.ilovepdf.com/compress_pdf');
      
      semestersItem.setChoices(semestersChoices);
      
      pages.forEach(function(page){
        page.setGoToPage(filePage);
      });
      
      data.sections.forEach(function(section){
        fileUploadItems.push(fileItem.duplicate().setTitle(section.code + ' - ' + section.name));
      });
      
      form.deleteItem(0);
      
      submissionFormId=form.getId();
      File(submissionFormId).moveToFolder(installationFolder);
      settings.set('SubmissionFormId', submissionFormId);
      settings.set('SubmissionFormUrl', form.getPublishedUrl());
      settings.set('SelectSemesterItemId', semestersItem.getId());
      settings.set('SelectCourseItemIds', 
                   serialize(selectCourseItems.map(
                     function(item){return item.getId();})));
      settings.set('FileUploadItemIds', 
                   serialize(fileUploadItems.map(
                     function(item){return item.getId();})));
      
      //Generate prefilled Form URLs for all courses
      data.semesters.forEach(function(semester,i){
        data.courses.filter(function(course){
          return course.semester==semester.code;
        }).forEach(function(course) {
          var response = form.createResponse()
          .withItemResponse(form.getItemById(semestersItem.getId()).asMultipleChoiceItem().createResponse(semester.name))
          .withItemResponse(form.getItemById(selectCourseItems[i].getId()).asListItem().createResponse(course.fullName));
          course.prefilledFormUrl = response.toPrefilledUrl(); 
        });
      });
      settings.set('CoursesWithUrls', serialize(data.courses));  
      
      //Initialize last timestamp to detect new submissions
      settings.set('LastTimestamp', new Date());
      
      return form;
    }
  }
  /****************************************************************/
  //Create Review Spreadsheet
  /****************************************************************/
  // One sheet for each semester
  // One row for each course
  // One column for each section
  function createReviewSheet(settings, installationFolder, data) {
    var reviewSheetFileId = settings.get('ReviewSheetFileId');
    if(!data.changed  && File(reviewSheetFileId).exists()) {
      Log.info(function(){return ['Install','An existing reviewSheetFileId was found and will be used.'];});
    } 
    else if( !data.changed  && installationFolder.getFileByName('Review Sheet').exists()) {
      reviewSheetFileId = installationFolder.getFileByName('Review Sheet').getId();
      settings.set('ReviewSheetFileId', reviewSheetFileId); 
      Log.info(function(){return ['Install','An existing Review Sheet file was found and will be used.'];});
    }
    else {
      //Create new review sheet
      reviewSheetFileId = SpreadsheetApp.create('Review Sheet').getId();
      File(reviewSheetFileId).moveToFolder(installationFolder);
      
      var reviewSheet = SpreadsheetApp.openById(reviewSheetFileId);
      var rule = SpreadsheetApp.newDataValidation().requireValueInList(
        ['Missing', 'N/A', 'Review','Send Comments','Awaiting Response',
         'Approved', 'Completed'], true).build();
      
      //Create a sheet for every semester
      data.semesters.forEach(function(semester,i){
        var sheet = reviewSheet.insertSheet(semester.name, i);
        var header = ['Code', 'Name', 'Coordinator'];
        data.sections.forEach(function(section){
          header.push('\'' + section.code);
          header.push(section.name);
        })
        sheet.appendRow(header);
        
        data.courses.forEach(function(course){
          if(course.semester==semester.code) {
            var row = [course.code,course.name, course.coordinator];
            data.sections.forEach(function(section) {
              row.push('');
              row.push('Missing');
            });
            sheet.appendRow(row);
          }
        });
        //Adding validation rule
        for(var i=0;i<data.sections.length;i++) 
          sheet.getRange(2, 5+2*i, sheet.getLastRow()-1, 1).setDataValidation(rule);
        
        sheet.autoResizeColumns(1, 3);
        sheet.setColumnWidths(4, sheet.getLastColumn()-3, 80);
        
        /********************************************************************
        * Add default sheet protection which warns editors before changing 
        * any cell except the Status column
        ********************************************************************/
        //TODO: Protect all sheet except status and coordinator
        /*"C2:C"+sheet.getLastRow()+
        ",E2:E"+sheet.getLastRow()+
        ",G2:G"+sheet.getLastRow()+
        ",I2:I"+sheet.getLastRow()+
        ",K2:K"+sheet.getLastRow()+
        ",M2:M"+sheet.getLastRow()+
        ",O2:O"+sheet.getLastRow()+
        ",Q2:Q"+sheet.getLastRow()+
        ",S2:S"+sheet.getLastRow()+
        ",U2:U"+sheet.getLastRow()+
        ",W2:W"+sheet.getLastRow() */
        var protection = sheet.protect();
        var editableRange = sheet.getRange("C2:W"+sheet.getLastRow());
        protection.setUnprotectedRanges([editableRange]);
        protection.setWarningOnly(true);
      });
      
      //Delete all other sheets from book
      reviewSheet.getSheets().forEach(function(sheet,i){
        if(i>=data.semesters.length)
          reviewSheet.deleteSheet(sheet);
      });
      
      settings.set('ReviewSheetFileId', reviewSheetFileId); 
      Log.info(function(){return ['Install','Review Sheet file was created.'];});
    }
  }    

  /****************************************************************/
  // Attach formSubmit trigger and periodic trigger
  /****************************************************************/
  function attachTriggers(form) {
    ScriptApp.getUserTriggers(form).forEach(function(trigger){
      ScriptApp.deleteTrigger(trigger);
    });
    
    ScriptApp.newTrigger('periodicTrigger')
    .timeBased()
    .everyHours(1)
    .create();
    
    ScriptApp.newTrigger("formSubmitTrigger")
    .forForm(form)
    .onFormSubmit()
    .create();
    
    Log.info(function(){return ['Install','Triggers attached successfully.'];});
  }  
  
  /****************************************************************/
  //Deploy as Web App and store Url
  /****************************************************************/
  function deployWebApp(settings) {
    var svc = ScriptApp.getService();
    if(svc.isEnabled()) {
      settings.set('webAppUrl', svc.getUrl());
      Log.info(function(){return ['Install','WebApp deployed successfully.'];});
    }
    else {
      Log.warn(function(){return ['Install','WebAppp wasn\'t published, please try to publish it manually, then resume installation.'];});
    }
  }
}

/********************************************************************/
/********************************************************************/
function renameCourse(code, newName) {
  renameCourseFolderContents(code, newName);
  renameCourseReviewSheets(code, newName);
  renameCourseSubmissionForm(code, newName);
  renameCourseSubmissionSheet(code,newName);
}

function renameCourseFolderContents(code, newName){
  var coursesFolder = DriveApp.getFolderById(config().coursesFolderId);
  var it = coursesFolder.getFolders();
  while(it.hasNext()) {
    var folder = it.next();
    if(folder.getName().indexOf(code)>=0) {
      renameCourseFolder(folder, code + ' ['+newName+']');

      var itt = folder.getFiles();
      while(itt.hasNext()) {
        var file=itt.next();
        renameCourseFile(file,newName);
      }
    }
  }
}

function renameCourseFolder(folder, newFolderName) {
  Logger.log('renameCourseFolder '+folder.getName()+' to ' + newFolderName );
  folder.setName(newFolderName);
}

function renameCourseFile(file,newName){
  var fi = splitStandardCourseFileName(file.getName());
  var newFileName = standardCourseFileName(fi.semester,fi.courseCode, fi.section, newName);         
  Logger.log('renameCourseFile('+file.getName()+' to ' + newFileName );
  file.setName(newFileName);
}

function renameCourseReviewSheets(code, newName){
  var reviewSheets = getReviewSheets();
  
  reviewSheets.forEach(function(sheet){
    var reviewRecords = getReviewRecords( sheet.file );

    reviewRecords.forEach(
      function (record) {
        if(record.courseCode==code) 
          record.renameCourse(newName);
      });
  });
}

function renameCourseSubmissionForm(code, newName){
  var form = FormApp.openById(config().submissionFormId);
  var items = form.getItems(FormApp.ItemType.LIST);
  items.forEach(function(item){
    var i = item.asListItem();
    i.setChoiceValues(
      i.getChoices().map(function(choice){
        if (choice.getValue().indexOf(code)>=0) {
          var newChoiceName = code+' ['+newName+']';
          Logger.log('renameCourseSubmissionForm '+choice.getValue()+' to '+newChoiceName)
          return newChoiceName ;
        }
        else
          return choice.getValue();
      }));
  });
}

function renameCourseSubmissionSheet(code,newName){
  var sheet = SpreadsheetApp.openById(config().submissionSpreadsheetId);
  for (var col=4; col<6; col++) {
    list = sheet.getSheetValues(2, col, 1000, 1);
    for(var row=1; row<list.length; row++) {
      if(list[row][0].indexOf(code)>=0) {
        var newChoiceName = code+' ['+newName+']';
        Logger.log('renameCourseSubmissionSheet '+list[row][0]+' to '+newChoiceName)
        list[row][0] = newChoiceName;
      }
      sheet.getSheets()[0].getRange(2, col, 1000, 1).setValues(list);
    }
  }
}

function renameElectricCircuitsCourse() {
  renameCourse('803324-3', 'Electric Circuits');
  
}

function renamePowerSystemProtectionCourse() {
  renameCourse('803532-3', 'Electric Power System Protection');
  
}

