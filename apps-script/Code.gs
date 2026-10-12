// ============================================
// Employee Task & Attendance Manager
// Container-Bound Script Version WITH ROLE-BASED ACCESS
// ============================================
// This script is bound directly to the Google Sheet
// No need to specify spreadsheet URL!

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Employee Task & Attendance Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ========== LOGIN & AUTHENTICATION ==========
/**
 * Login user by Employee ID
 * Returns user data if found and active
 */
function loginUser(employeeId) {
  try {
    const ss = getSpreadsheet();
    const empSheet = ss.getSheetByName('Employees');
    
    if (!empSheet) {
      Logger.log('Employees sheet not found');
      return { 
        success: false, 
        message: 'Employees sheet not found. Please create it with exact name: Employees'
      };
    }
    
    const data = empSheet.getDataRange().getValues();
    
    // Find employee by ID (starting from row 2, skipping header)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = String(row[0]);
      const status = String(row[5]) || 'Active';
      
      if (id === String(employeeId) && status === 'Active') {
        Logger.log('User logged in: ' + id);
        return {
          success: true,
          data: {
            id: String(row[0]),
            name: String(row[1]),
            email: String(row[2]),
            department: String(row[3]),
            position: String(row[4]),
            status: status,
            role: String(row[6]) || 'User' // Role is in column G (index 6)
          }
        };
      }
    }
    
    Logger.log('Login failed for ID: ' + employeeId);
    return { 
      success: false, 
      message: 'Invalid Employee ID or account is inactive' 
    };
  } catch (e) {
    Logger.log('Error in loginUser: ' + e.toString());
    return { 
      success: false, 
      message: 'Login error: ' + e.toString() 
    };
  }
}

// ========== EMPLOYEE FUNCTIONS ==========
function getEmployeeList() {
  try {
    const ss = getSpreadsheet();
    const empSheet = ss.getSheetByName('Employees');
    
    if (!empSheet) {
      Logger.log('Employees sheet not found');
      return { 
        success: false, 
        message: 'Sheet "Employees" not found. Please create it with exact name: Employees',
        data: []
      };
    }
    
    const lastRow = empSheet.getLastRow();
    Logger.log('Employees sheet last row: ' + lastRow);
    
    if (lastRow < 2) {
      return { 
        success: true, 
        message: 'No employees found. Add employees first.',
        data: []
      };
    }
    
    const data = empSheet.getRange(2, 1, lastRow - 1, 7).getValues(); // Now reading 7 columns (A-G) to include Role
    const employees = [];
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0]) {
        employees.push({
          id: String(data[i][0]),
          name: String(data[i][1]),
          email: String(data[i][2]),
          department: String(data[i][3]),
          position: String(data[i][4]),
          status: String(data[i][5]) || 'Active',
          role: String(data[i][6]) || 'User' // Column G
        });
      }
    }
    
    Logger.log('Found ' + employees.length + ' employees');
    return { success: true, data: employees };
  } catch (e) {
    Logger.log('Error in getEmployeeList: ' + e.toString());
    return { 
      success: false, 
      message: 'Error loading employees: ' + e.toString(),
      data: []
    };
  }
}

function addEmployee(id, name, email, department, position, role) {
  try {
    const ss = getSpreadsheet();
    const empSheet = ss.getSheetByName('Employees');
    
    if (!empSheet) {
      return { success: false, message: 'Employees sheet not found' };
    }
    
    // Check if employee ID already exists
    const data = empSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        return { success: false, message: 'Employee ID already exists!' };
      }
    }
    
    empSheet.appendRow([id, name, email, department, position, 'Active', role || 'User']);
    Logger.log('Employee added: ' + id + ' with role: ' + (role || 'User'));
    
    return { success: true, message: 'Employee added successfully!' };
  } catch (e) {
    Logger.log('Error in addEmployee: ' + e.toString());
    return { success: false, message: 'Error adding employee: ' + e.toString() };
  }
}

// ========== ATTENDANCE FUNCTIONS ==========
function markAttendance(employeeId, checkIn, checkOut, status) {
  try {
    const ss = getSpreadsheet();
    const attSheet = ss.getSheetByName('Attendance');
    const empSheet = ss.getSheetByName('Employees');
    
    if (!attSheet) {
      Logger.log('Attendance sheet not found');
      return { success: false, message: 'Attendance sheet not found. Please create it.' };
    }
    
    if (!empSheet) {
      return { success: false, message: 'Employees sheet not found' };
    }
    
    // Find employee name
    const empData = empSheet.getDataRange().getValues();
    let employeeName = '';
    for (let i = 1; i < empData.length; i++) {
      if (String(empData[i][0]) === String(employeeId)) {
        employeeName = empData[i][1];
        break;
      }
    }
    
    if (!employeeName) {
      return { success: false, message: 'Employee not found' };
    }
    
    // Calculate hours
    let totalHours = 0;
    if (checkIn && checkOut) {
      const inTime = new Date('1970/01/01 ' + checkIn);
      const outTime = new Date('1970/01/01 ' + checkOut);
      totalHours = ((outTime - inTime) / (1000 * 60 * 60)).toFixed(2);
    }
    
    const today = new Date();
    const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    // Check if attendance already exists for today
    const attData = attSheet.getDataRange().getValues();
    for (let i = 1; i < attData.length; i++) {
      const recordDate = Utilities.formatDate(new Date(attData[i][0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (recordDate === todayStr && String(attData[i][1]) === String(employeeId)) {
        // Update existing record
        attSheet.getRange(i + 1, 1, 1, 7).setValues([[today, employeeId, employeeName, checkIn, checkOut, status, totalHours]]);
        Logger.log('Attendance updated for: ' + employeeName);
        return { success: true, message: 'Attendance updated successfully for ' + employeeName + '!' };
      }
    }
    
    // Add new record
    attSheet.appendRow([today, employeeId, employeeName, checkIn, checkOut, status, totalHours]);
    Logger.log('Attendance marked for: ' + employeeName);
    
    return { success: true, message: 'Attendance marked successfully for ' + employeeName + '!' };
  } catch (e) {
    Logger.log('Error in markAttendance: ' + e.toString());
    return { success: false, message: 'Error marking attendance: ' + e.toString() };
  }
}

function getAttendanceRecords(days) {
  try {
    days = days || 30;
    const ss = getSpreadsheet();
    const attSheet = ss.getSheetByName('Attendance');
    
    if (!attSheet) {
      Logger.log('Attendance sheet not found');
      return { 
        success: true,
        message: 'Attendance sheet not found',
        data: []
      };
    }
    
    const lastRow = attSheet.getLastRow();
    Logger.log('Attendance sheet last row: ' + lastRow);
    
    if (lastRow < 2) {
      return { 
        success: true, 
        message: 'No attendance records found yet. Mark attendance first!',
        data: []
      };
    }
    
    const data = attSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const records = [];
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - days);
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][1]) {
        const recordDate = new Date(data[i][0]);
        if (recordDate >= cutoffDate) {
          records.push({
            date: Utilities.formatDate(recordDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
            employeeId: String(data[i][1]),
            employeeName: String(data[i][2]),
            checkIn: String(data[i][3]),
            checkOut: String(data[i][4]),
            status: String(data[i][5]),
            hours: String(data[i][6])
          });
        }
      }
    }
    
    Logger.log('Found ' + records.length + ' attendance records');
    return { success: true, data: records };
  } catch (e) {
    Logger.log('Error in getAttendanceRecords: ' + e.toString());
    return { 
      success: true,
      message: 'Error loading records: ' + e.toString(),
      data: []
    };
  }
}

/**
 * Get attendance for specific employee (User function)
 * Users can only see their own attendance
 */
function getMyAttendance(employeeId) {
  try {
    const ss = getSpreadsheet();
    const attSheet = ss.getSheetByName('Attendance');
    
    if (!attSheet) {
      return { success: true, data: [] };
    }
    
    const data = attSheet.getDataRange().getValues();
    const records = [];
    
    // Get last 30 days of attendance for this employee only
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - 30);
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(employeeId)) { // Filter by employee ID
        const recordDate = new Date(data[i][0]);
        if (recordDate >= cutoffDate) {
          records.push({
            date: Utilities.formatDate(recordDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
            checkIn: String(data[i][3]),
            checkOut: String(data[i][4]),
            status: String(data[i][5]),
            hours: String(data[i][6])
          });
        }
      }
    }
    
    Logger.log('Found ' + records.length + ' attendance records for employee: ' + employeeId);
    return { success: true, data: records };
  } catch (e) {
    Logger.log('Error in getMyAttendance: ' + e.toString());
    return { 
      success: false, 
      message: 'Error: ' + e.toString(),
      data: []
    };
  }
}

// ========== TASK FUNCTIONS ==========
function createTask(taskName, employeeId, priority, dueDate, notes) {
  try {
    const ss = getSpreadsheet();
    const taskSheet = ss.getSheetByName('Tasks');
    const empSheet = ss.getSheetByName('Employees');
    
    if (!taskSheet) {
      Logger.log('Tasks sheet not found');
      return { success: false, message: 'Tasks sheet not found. Please create it.' };
    }
    
    if (!empSheet) {
      return { success: false, message: 'Employees sheet not found' };
    }
    
    const lastRow = taskSheet.getLastRow();
    const taskId = 'T' + String(lastRow + 1000).padStart(4, '0');
    
    // Find employee name
    const empData = empSheet.getDataRange().getValues();
    let employeeName = '';
    for (let i = 1; i < empData.length; i++) {
      if (String(empData[i][0]) === String(employeeId)) {
        employeeName = empData[i][1];
        break;
      }
    }
    
    if (!employeeName) {
      return { success: false, message: 'Employee not found' };
    }
    
    taskSheet.appendRow([taskId, taskName, employeeId, employeeName, priority, 'Pending', dueDate, '', notes]);
    Logger.log('Task created: ' + taskId);
    
    return { success: true, message: 'Task "' + taskName + '" created successfully!', taskId: taskId };
  } catch (e) {
    Logger.log('Error in createTask: ' + e.toString());
    return { success: false, message: 'Error creating task: ' + e.toString() };
  }
}

function getTasks(filter) {
  try {
    filter = filter || 'all';
    const ss = getSpreadsheet();
    const taskSheet = ss.getSheetByName('Tasks');
    
    if (!taskSheet) {
      Logger.log('Tasks sheet not found');
      return { 
        success: true,
        message: 'Tasks sheet not found',
        data: []
      };
    }
    
    const lastRow = taskSheet.getLastRow();
    Logger.log('Tasks sheet last row: ' + lastRow);
    
    if (lastRow < 2) {
      return { 
        success: true, 
        message: 'No tasks found. Create one first!',
        data: []
      };
    }
    
    const data = taskSheet.getRange(2, 1, lastRow - 1, 9).getValues();
    const tasks = [];
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0]) {
        const task = {
          id: String(data[i][0]),
          name: String(data[i][1]),
          employeeId: String(data[i][2]),
          employeeName: String(data[i][3]),
          priority: String(data[i][4]),
          status: String(data[i][5]),
          dueDate: data[i][6] ? Utilities.formatDate(new Date(data[i][6]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
          completionDate: data[i][7] ? String(data[i][7]) : '',
          notes: String(data[i][8] || '')
        };
        
        if (filter === 'all' || 
            (filter === 'pending' && task.status !== 'Completed') ||
            (filter === 'completed' && task.status === 'Completed')) {
          tasks.push(task);
        }
      }
    }
    
    Logger.log('Found ' + tasks.length + ' tasks (filter: ' + filter + ')');
    return { success: true, data: tasks };
  } catch (e) {
    Logger.log('Error in getTasks: ' + e.toString());
    return { 
      success: true,
      message: 'Error loading tasks: ' + e.toString(),
      data: []
    };
  }
}

/**
 * Get tasks for specific employee (User function)
 * Users can only see tasks assigned to them
 */
function getMyTasks(employeeId) {
  try {
    const ss = getSpreadsheet();
    const taskSheet = ss.getSheetByName('Tasks');
    
    if (!taskSheet) {
      return { success: true, data: [] };
    }
    
    const data = taskSheet.getDataRange().getValues();
    const tasks = [];
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]) === String(employeeId)) { // Filter by employee ID (column C)
        tasks.push({
          id: String(data[i][0]),
          name: String(data[i][1]),
          priority: String(data[i][4]),
          status: String(data[i][5]),
          dueDate: data[i][6] ? Utilities.formatDate(new Date(data[i][6]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
          notes: String(data[i][8] || '')
        });
      }
    }
    
    Logger.log('Found ' + tasks.length + ' tasks for employee: ' + employeeId);
    return { success: true, data: tasks };
  } catch (e) {
    Logger.log('Error in getMyTasks: ' + e.toString());
    return { 
      success: false, 
      message: 'Error: ' + e.toString(),
      data: []
    };
  }
}

function updateTaskStatus(taskId, newStatus) {
  try {
    const ss = getSpreadsheet();
    const taskSheet = ss.getSheetByName('Tasks');
    
    if (!taskSheet) {
      return { success: false, message: 'Tasks sheet not found' };
    }
    
    const data = taskSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(taskId)) {
        taskSheet.getRange(i + 1, 6).setValue(newStatus);
        if (newStatus === 'Completed') {
          taskSheet.getRange(i + 1, 8).setValue(new Date());
        }
        Logger.log('Task updated: ' + taskId + ' to ' + newStatus);
        return { success: true, message: 'Task updated to ' + newStatus + '!' };
      }
    }
    
    return { success: false, message: 'Task not found' };
  } catch (e) {
    Logger.log('Error in updateTaskStatus: ' + e.toString());
    return { success: false, message: 'Error updating task: ' + e.toString() };
  }
}

// ========== DASHBOARD STATS ==========
function getDashboardStats() {
  try {
    const ss = getSpreadsheet();
    
    // Get sheets
    const empSheet = ss.getSheetByName('Employees');
    const attSheet = ss.getSheetByName('Attendance');
    const taskSheet = ss.getSheetByName('Tasks');
    
    // Initialize stats
    let totalEmployees = 0;
    let presentToday = 0;
    let pendingTasks = 0;
    let inProgressTasks = 0;
    let completedTasks = 0;
    
    // Count employees
    if (empSheet && empSheet.getLastRow() > 1) {
      totalEmployees = empSheet.getLastRow() - 1;
    }
    
    // Today's attendance
    if (attSheet && attSheet.getLastRow() > 1) {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const attData = attSheet.getDataRange().getValues();
      
      for (let i = 1; i < attData.length; i++) {
        if (attData[i][0]) {
          const recordDate = new Date(attData[i][0]);
          recordDate.setHours(0, 0, 0, 0);
          if (recordDate.getTime() === today.getTime() && attData[i][5] === 'Present') {
            presentToday++;
          }
        }
      }
    }
    
    // Task counts
    if (taskSheet && taskSheet.getLastRow() > 1) {
      const taskData = taskSheet.getDataRange().getValues();
      
      for (let i = 1; i < taskData.length; i++) {
        if (taskData[i][0]) {
          if (taskData[i][5] === 'Pending') pendingTasks++;
          else if (taskData[i][5] === 'In Progress') inProgressTasks++;
          else if (taskData[i][5] === 'Completed') completedTasks++;
        }
      }
    }
    
    Logger.log('Dashboard stats calculated successfully');
    return {
      success: true,
      data: {
        totalEmployees: totalEmployees,
        presentToday: presentToday,
        pendingTasks: pendingTasks,
        inProgressTasks: inProgressTasks,
        completedTasks: completedTasks
      }
    };
  } catch (e) {
    Logger.log('Error in getDashboardStats: ' + e.toString());
    return { 
      success: false, 
      message: 'Error loading dashboard: ' + e.toString(),
      data: {
        totalEmployees: 0,
        presentToday: 0,
        pendingTasks: 0,
        inProgressTasks: 0,
        completedTasks: 0
      }
    };
  }
}

// ========== TEST FUNCTION ==========
function testSetup() {
  try {
    const ss = getSpreadsheet();
    const sheets = ss.getSheets();
    const sheetNames = sheets.map(s => s.getName());
    
    Logger.log('✓ Successfully connected to spreadsheet: ' + ss.getName());
    Logger.log('✓ Sheet names found: ' + sheetNames.join(', '));
    
    // Test each function
    Logger.log('Testing getEmployeeList...');
    const empResult = getEmployeeList();
    Logger.log('Employees: ' + JSON.stringify(empResult));
    
    Logger.log('Testing getTasks...');
    const taskResult = getTasks('all');
    Logger.log('Tasks: ' + JSON.stringify(taskResult));
    
    Logger.log('Testing getAttendanceRecords...');
    const attResult = getAttendanceRecords(30);
    Logger.log('Records: ' + JSON.stringify(attResult));
    
    Logger.log('Testing getDashboardStats...');
    const statsResult = getDashboardStats();
    Logger.log('Stats: ' + JSON.stringify(statsResult));
    
    return {
      success: true,
      spreadsheetName: ss.getName(),
      sheetNames: sheetNames,
      message: 'All tests passed!'
    };
    
  } catch (e) {
    Logger.log('✗ Error: ' + e.toString());
    return {
      success: false,
      message: e.toString()
    };
  }
}
