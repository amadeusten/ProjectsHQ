/**
 * PROJECT SETUP - Creation and Folder Management
 * Handles new project creation, folder structure, and asset tracker setup
 */

// ============================================================================
// MAIN PROJECT CREATION FUNCTION
// ============================================================================

function createNewProject(setupData) {
  try {
    // Validate required fields
    const required = ['projectNumber', 'projectName', 'client', 'description', 'emailNotifications', 'projectManagers'];
    for (const field of required) {
      if (!setupData[field] || setupData[field].toString().trim() === '') {
        return { 
          success: false, 
          message: `Required field missing: ${field}`
        };
      }
    }
    
    // Get HQ configuration
    const hqConfig = getHQConfig();
    const parentFolderId = hqConfig.PARENT_FOLDER_ID;
    const defaultEmail = hqConfig.DEFAULT_EMAIL;
    
    if (!parentFolderId) {
      return { 
        success: false, 
        message: 'Parent folder not configured. Check HQ Config sheet.'
      };
    }
    
    // Combine emails
    const projectEmails = setupData.emailNotifications.trim();
    const combinedEmails = defaultEmail + (projectEmails ? ', ' + projectEmails : '');
    setupData.alertEmails = combinedEmails;
    
    // Create folder structure
    const folderResult = createProjectFolderStructure(setupData, parentFolderId);
    if (!folderResult.success) {
      return folderResult;
    }
    
    // Create Asset Tracker
    const trackerResult = createAssetTracker(setupData, folderResult);
    if (!trackerResult.success) {
      return trackerResult;
    }
    
    // Add to Projects List sheet
    const projectSheetData = {
      projectNumber: setupData.projectNumber,
      projectName: setupData.projectName,
      client: setupData.client,
      description: setupData.description,
      projectManagers: Array.isArray(setupData.projectManagers) 
        ? setupData.projectManagers.join(', ') 
        : setupData.projectManagers,
      emailNotifications: combinedEmails,
      startDate: setupData.startDate,
      inHandsDate: setupData.inHandsDate,
      assetTrackerUrl: trackerResult.trackerUrl
    };
    
    addProjectToSheet(projectSheetData);
    
    return {
      success: true,
      message: 'Project created successfully!',
      projectNumber: setupData.projectNumber,
      assetTrackerUrl: trackerResult.trackerUrl
    };
    
  } catch (error) {
    console.error('Error creating project:', error);
    return { 
      success: false, 
      message: 'Error: ' + error.toString()
    };
  }
}

// ============================================================================
// FOLDER STRUCTURE CREATION
// ============================================================================

function createProjectFolderStructure(setupData, parentFolderId) {
  try {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    
    // Create main project folder
    const projectFolderName = setupData.projectNumber;
    const projectFolder = parentFolder.createFolder(projectFolderName);
    const projectFolderId = projectFolder.getId();
    
    // Create subfolder structure
    createSubfolders(projectFolder, setupData.projectName);
    
    return {
      success: true,
      projectFolder: projectFolder,
      projectFolderId: projectFolderId
    };
    
  } catch (error) {
    console.error('Error creating folder structure:', error);
    return { 
      success: false, 
      message: `Folder creation failed: ${error.toString()}`
    };
  }
}

function createSubfolders(projectFolder, projectName) {
  // Level 1 folders
  projectFolder.createFolder('01 - Admin Docs');
  const productionFiles = projectFolder.createFolder('02 - Production Files');
  const projectFiles = projectFolder.createFolder('03 - Project Files');
  projectFolder.createFolder('04 - Vendor Docs');
  
  // Level 2 folders under 02 - Production Files
  productionFiles.createFolder(`0 - ${projectName} Artwork Files`);
  productionFiles.createFolder('On Hold');
  
  // Level 2 folders under 03 - Project Files
  projectFiles.createFolder('Sections');
  projectFiles.createFolder(`Team Docs - ${projectName}`);
}

// ============================================================================
// ASSET TRACKER CREATION
// ============================================================================

function createAssetTracker(setupData, folderResult) {
  try {
    // Create new spreadsheet with proper OAuth scope
    const trackerName = `Production - ${setupData.projectNumber}`;
    const tracker = SpreadsheetApp.create(trackerName);
    const trackerId = tracker.getId();
    const trackerUrl = tracker.getUrl();
    
    // Move to project folder
    const trackerFile = DriveApp.getFileById(trackerId);
    trackerFile.moveTo(folderResult.projectFolder);
    
    // Set up the tracker
    setupAssetTrackerSheets(tracker, setupData, folderResult);
    
    return {
      success: true,
      trackerId: trackerId,
      trackerUrl: trackerUrl
    };
    
  } catch (error) {
    console.error('Error creating asset tracker:', error);
    return {
      success: false,
      message: `Asset tracker creation failed: ${error.toString()}`
    };
  }
}

function setupAssetTrackerSheets(tracker, setupData, folderResult) {
  // Rename first sheet to Master
  const mainSheet = tracker.getSheets()[0];
  mainSheet.setName(`${setupData.projectNumber} Master`);
  
  // Create headers
  const headers = [
    'ID', 'Area', 'Asset', 'Status', 'Dimensions', 'Quantity', 'Item', 'Material',
    'Due Date', 'Strike Date', 'Venue', 'Location', 'Artwork', 'Image Link',
    'Double Sided', 'Diecut', 'Production Status', 'Edit'
  ];
  
  mainSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  mainSheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  mainSheet.setFrozenRows(1);
  
  // Set column widths
  mainSheet.setColumnWidth(1, 80);   // ID
  mainSheet.setColumnWidth(2, 100);  // Area
  mainSheet.setColumnWidth(3, 200);  // Asset
  mainSheet.setColumnWidth(4, 120);  // Status
  mainSheet.setColumnWidth(5, 120);  // Dimensions
  mainSheet.setColumnWidth(13, 200); // Artwork
  mainSheet.setColumnWidth(14, 200); // Image Link
  
  // Create ProjectConfig sheet (HIDDEN)
  createProjectConfigSheet(tracker, setupData, folderResult);
  
  // Create MaterialIDMap sheet (starts empty)
  const materialSheet = tracker.insertSheet('MaterialIDMap');
  materialSheet.getRange(1, 1, 1, 2).setValues([['Material', 'ID Prefix']]);
  materialSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e8f0fe');
  materialSheet.hideSheet();
  
  // Create AssetLog sheet
  const logSheet = tracker.insertSheet('AssetLog');
  logSheet.getRange(1, 1, 1, 4).setValues([['LogID', 'ProjectRow', 'Timestamp', 'FormData']]);
  logSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e8f0fe');
  logSheet.hideSheet();
  
  // Create dropdown lists with initial values
  createDropdownSheets(tracker);
  
  // Install complete Asset Tracker code
  installAssetTrackerCode(tracker, setupData, folderResult);
  
  SpreadsheetApp.flush();
}

function createProjectConfigSheet(tracker, setupData, folderResult) {
  const configSheet = tracker.insertSheet('ProjectConfig', 0);
  
  const configData = [
    ['PROJECT CONFIGURATION', 'Value', 'Description'],
    ['', '', ''],
    ['Basic Information', '', ''],
    ['Project Name', setupData.projectName, 'Full name of the project'],
    ['Project Code', setupData.projectNumber, 'Project number/code'],
    ['Client Name', setupData.client, 'Client or company name'],
    ['', '', ''],
    ['Email Alerts', '', ''],
    ['Alert Recipients', setupData.alertEmails, 'Email addresses for alerts'],
    ['', '', ''],
    ['Google Drive Folders', '', ''],
    ['Main Folder ID', folderResult.projectFolderId, 'Google Drive folder ID'],
    ['Artwork Folder', `0 - ${setupData.projectName} Artwork Files`, 'Folder for artwork files'],
    ['Production Folder', '02 - Production Files', 'Production files folder'],
    ['Team Docs Folder', `Team Docs - ${setupData.projectName}`, 'Team documentation folder'],
    ['', '', ''],
    ['Sheet Configuration', '', ''],
    ['Master Sheet Name', `${setupData.projectNumber} Master`, 'Main asset tracking sheet'],
    ['Log Sheet Name', 'AssetLog', 'Change log sheet'],
    ['Material ID Map Sheet', 'MaterialIDMap', 'Material ID mapping sheet']
  ];
  
  configSheet.getRange(1, 1, configData.length, 3).setValues(configData);
  
  configSheet.setColumnWidth(1, 200);
  configSheet.setColumnWidth(2, 300);
  configSheet.setColumnWidth(3, 350);
  
  configSheet.getRange('A1:C1')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12);
  
  const sectionRows = [3, 8, 11, 17];
  sectionRows.forEach(row => {
    configSheet.getRange(row, 1, 1, 3).setBackground('#e8f0fe').setFontWeight('bold');
  });
  
  configSheet.getRange('B2:B30').setBackground('#f8f9fa');
  
  // HIDE the ProjectConfig sheet
  configSheet.hideSheet();
}

function createDropdownSheets(tracker) {
  const dropdownData = {
    'ItemsList': ['Banner', 'Sign', 'Poster', 'Decal', 'Display', 'Billboard'],
    'MaterialsList': ['Adhesive Vinyl - Matte', 'Foamcore - 1/4"', 'Foamcore - 1/2"', 'Gatorplast - 1/4"', 'Gatorplast - 1/2"'],
    'StatusesList': ['New Asset', 'In Progress', 'Awaiting Approval', 'Approved', 'In Production', 'Delivered', 'On Hold', 'Requires Attn'],
    'VenuesList': ['Main Hall', 'Conference Room', 'Lobby', 'Outdoor'],
    'AreasList': ['North', 'South', 'East', 'West', 'Central'],
    'ProductionStatusesList': ['Processing', 'Printing', 'Cutting', 'Finishing', 'Ready', 'Picked Up']
  };
  
  Object.entries(dropdownData).forEach(([sheetName, values]) => {
    const sheet = tracker.insertSheet(sheetName);
    sheet.getRange(1, 1).setValue(sheetName.replace('List', ''));
    sheet.getRange(1, 1).setFontWeight('bold').setBackground('#e8f0fe');
    
    if (values && values.length > 0) {
      const data = values.map(v => [v]);
      sheet.getRange(2, 1, data.length, 1).setValues(data);
    }
    
    sheet.hideSheet();
  });
}

function installAssetTrackerCode(tracker, setupData, folderResult) {
  try {
    const scriptProject = ScriptApp.create(`${setupData.projectName} - Asset Tracker Code`);
    const projectId = scriptProject.getId();
    
    // Bind the script to the spreadsheet
    const boundScript = tracker.getUrl().match(/\/d\/([a-zA-Z0-9-_]+)/)[1];
    
    // Note: We cannot programmatically create bound scripts from Apps Script
    // Instead, we'll add installation instructions to the tracker
    
    const instructionSheet = tracker.insertSheet('📌 SETUP INSTRUCTIONS');
    
    const instructions = [
      ['ASSET TRACKER SETUP - COPY CODE FILES'],
      [''],
      ['This Asset Tracker is ready to use, but needs code installed to function fully.'],
      [''],
      ['📋 REQUIRED STEPS:'],
      [''],
      ['1. Open Extensions → Apps Script in this spreadsheet'],
      ['2. Copy the following files from Project HQ Apps Script:'],
      ['   • Code.gs (Asset Tracker version)'],
      ['   • AssetForm.html'],
      ['   • HXCommentDialog.html'],
      ['   • DropdownEditor.html'],
      ['   • ReorderAssetForm.html'],
      ['   • appsscript.json'],
      [''],
      ['3. Update these values in Code.gs:'],
      [`   • CONFIG.MASTER_SHEET_NAME = '${setupData.projectNumber} Master'`],
      [`   • CONFIG.PROJECT_CODE = '${setupData.projectNumber}'`],
      [`   • CONFIG.DRIVE_FOLDER_ID = '${folderResult.projectFolderId}'`],
      [`   • Menu name = '${setupData.projectName}'`],
      [`   • Email recipients = '${setupData.alertEmails}'`],
      [''],
      ['4. Save and refresh this spreadsheet'],
      ['5. You will see a menu with the project name appear'],
      ['6. Delete this instruction sheet'],
      [''],
      ['✅ Once complete, you can add assets, manage dropdowns, and use all features!'],
      [''],
      ['Need help? Check the Project HQ documentation.']
    ];
    
    instructionSheet.getRange(1, 1, instructions.length, 1).setValues(instructions);
    instructionSheet.setColumnWidth(1, 700);
    instructionSheet.getRange('A1').setFontSize(14).setFontWeight('bold').setBackground('#fef3c7');
    instructionSheet.getRange('A5').setFontWeight('bold').setFontSize(12);
    instructionSheet.getRange('A3').setFontColor('#ef4444');
    
    // Make it the active sheet
    tracker.setActiveSheet(instructionSheet);
    
  } catch (error) {
    console.error('Error in installAssetTrackerCode:', error);
  }
}
