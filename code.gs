// ตั้งค่า Google Sheets - ใช้ชีทเดียว 4 แท็บ
const MAIN_SHEET_ID = '1sptogt8HQOlkdb6Ml0j2RFLI4I3ePfF6KnLgXM2uWho'; // Sheet หลัก

// แท็บต่างๆ ในชีทเดียว
const SHEET_TABS = {
  users: 'Users',
  server1: 'Server1_Files',
  server2: 'Server2_Files',
  server3: 'Server3_Files'
};

// Drive Folders แยกตาม Server
const SERVER_FOLDERS = {
  server1: '1a5Neo6K80TsBv8LFtU2tpQlwqnxJKONW',
  server2: '1kvZHHbQFYn0aGmLfa4z1TAYvjKK928vz',
  server3: '16FeoivQugdXOywNv8Cg__dygYKRac_HP'
};

// ฟังก์ชันสำหรับแสดงหน้าเว็บ
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ฟังก์ชันเริ่มต้น - สร้างข้อมูลเริ่มต้น
function initializeSystem() {
  try {
    // สร้าง Users Tab
    const usersSheet = getOrCreateSheet(SHEET_TABS.users);
    if (usersSheet.getLastRow() <= 1) {
      // สร้าง Header
      usersSheet.getRange(1, 1, 1, 5).setValues([['Username', 'Password', 'Role', 'Status', 'Created']]);
      
      // เพิ่มผู้ใช้เริ่มต้น
      const defaultUsers = [
        ['admin', 'admin123', 'admin', 'active', new Date()],
        ['user1', 'user123', 'user', 'active', new Date()]
      ];
      usersSheet.getRange(2, 1, defaultUsers.length, 5).setValues(defaultUsers);
    }
    
    // สร้าง Server Tabs
    Object.keys(SERVER_FOLDERS).forEach(server => {
      const serverSheet = getOrCreateSheet(SHEET_TABS[server]);
      if (serverSheet.getLastRow() <= 1) {
        // สร้าง Header
        serverSheet.getRange(1, 1, 1, 8).setValues([['FileID', 'Filename', 'Owner', 'Size', 'MimeType', 'UploadDate', 'DriveFileId', 'Server']]);
      }
    });
    
    // สร้าง Drive Folders
    Object.keys(SERVER_FOLDERS).forEach(server => {
      createDriveFolder(server);
    });
    
    return { success: true, message: 'ระบบเริ่มต้นสำเร็จ' };
  } catch (error) {
    console.error('Initialize error:', error);
    return { success: false, message: 'เกิดข้อผิดพลาดในการเริ่มต้นระบบ: ' + error.toString() };
  }
}

// ฟังก์ชันสำหรับสร้างหรือเปิด Sheet Tab
function getOrCreateSheet(tabName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(MAIN_SHEET_ID);
    let sheet = spreadsheet.getSheetByName(tabName);
    
    if (!sheet) {
      // สร้างแท็บใหม่ถ้าไม่มี
      sheet = spreadsheet.insertSheet(tabName);
      console.log('Created new sheet tab:', tabName);
    }
    
    return sheet;
  } catch (error) {
    console.error('Error creating/opening sheet tab:', error);
    throw new Error('ไม่สามารถเปิดหรือสร้าง Google Sheets แท็บได้');
  }
}

// ฟังก์ชันสำหรับสร้าง Drive Folder
function createDriveFolder(server) {
  try {
    const folderId = SERVER_FOLDERS[server];
    if (folderId && !folderId.startsWith('YOUR_')) {
      return DriveApp.getFolderById(folderId);
    } else {
      const folder = DriveApp.createFolder(`File Management System - ${server.toUpperCase()}`);
      console.log(`Created new folder for ${server}:`, folder.getId());
      return folder;
    }
  } catch (error) {
    console.error(`Error creating/opening folder for ${server}:`, error);
    throw new Error(`ไม่สามารถเปิดหรือสร้าง Google Drive Folder สำหรับ ${server} ได้`);
  }
}

// ฟังก์ชันล็อกอิน
function login(username, password) {
  try {
    const usersSheet = getOrCreateSheet(SHEET_TABS.users);
    const data = usersSheet.getDataRange().getValues();
    
    // หาผู้ใช้
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === username && row[1] === password && row[3] === 'active') {
        return {
          success: true,
          user: {
            username: row[0],
            role: row[2]
          }
        };
      }
    }
    
    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  } catch (error) {
    console.error('Login error:', error);
    return { success: false, message: 'เกิดข้อผิดพลาดในการล็อกอิน: ' + error.toString() };
  }
}

// ฟังก์ชันโหลดข้อมูล Dashboard
function getDashboardData(username, role) {
  try {
    const usersSheet = getOrCreateSheet(SHEET_TABS.users);
    const usersData = usersSheet.getDataRange().getValues();
    
    let allFiles = [];
    let serverStats = {
      server1Files: 0,
      server2Files: 0,
      server3Files: 0
    };
    
    // รวบรวมข้อมูลจากทุก Server Tab
    Object.keys(SERVER_FOLDERS).forEach(server => {
      try {
        const serverSheet = getOrCreateSheet(SHEET_TABS[server]);
        const filesData = serverSheet.getDataRange().getValues();
        
        // นับไฟล์ใน server นี้
        const serverFileCount = filesData.length - 1; // ลบ header
        serverStats[`${server}Files`] = Math.max(0, serverFileCount);
        
        // เพิ่มไฟล์ในรายการ
        for (let i = 1; i < filesData.length; i++) {
          const row = filesData[i];
          if (role === 'admin' || row[2] === username) {
            allFiles.push({
              fileId: row[0],
              filename: row[1],
              owner: row[2],
              size: formatFileSize(row[3]),
              uploadDate: formatDate(row[5]),
              driveFileId: row[6],
              server: server
            });
          }
        }
      } catch (error) {
        console.error(`Error loading data from ${server}:`, error);
        serverStats[`${server}Files`] = 0;
      }
    });
    
    // นับจำนวนผู้ใช้
    let totalUsers = usersData.length - 1; // ลบ header
    
    return {
      success: true,
      data: {
        server1Files: serverStats.server1Files,
        server2Files: serverStats.server2Files,
        server3Files: serverStats.server3Files,
        totalUsers: totalUsers,
        files: allFiles.reverse() // แสดงไฟล์ล่าสุดก่อน
      }
    };
  } catch (error) {
    console.error('Dashboard data error:', error);
    return { success: false, message: 'เกิดข้อผิดพลาดในการโหลดข้อมูล: ' + error.toString() };
  }
}

// ฟังก์ชันอัปโหลดไฟล์
function uploadFile(filename, fileData, owner, mimeType, server) {
  try {
    // ตรวจสอบ server
    if (!SHEET_TABS[server] || !SERVER_FOLDERS[server]) {
      return { success: false, message: 'เซิร์ฟเวอร์ไม่ถูกต้อง' };
    }
    
    // แปลง base64 เป็น blob
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, filename);
    
    // อัปโหลดไปยัง Drive ของ server ที่เลือก
    const folder = createDriveFolder(server);
    const driveFile = folder.createFile(blob);
    
    // บันทึกข้อมูลใน Tab ของ server ที่เลือก
    const serverSheet = getOrCreateSheet(SHEET_TABS[server]);
    const fileId = Utilities.getUuid();
    
    const newRow = [
      fileId,
      filename,
      owner,
      blob.getBytes().length,
      mimeType,
      new Date(),
      driveFile.getId(),
      server
    ];
    
    serverSheet.appendRow(newRow);
    
    return { success: true, message: `อัปโหลดไฟล์ไปยัง ${server.toUpperCase()} สำเร็จ` };
  } catch (error) {
    console.error('Upload error:', error);
    return { success: false, message: 'เกิดข้อผิดพลาดในการอัปโหลด: ' + error.toString() };
  }
}

// ฟังก์ชันดาวน์โหลดไฟล์
function downloadFile(fileId) {
  try {
    // ค้นหาไฟล์ในทุก server tab
    for (const server of Object.keys(SERVER_FOLDERS)) {
      try {
        const serverSheet = getOrCreateSheet(SHEET_TABS[server]);
        const data = serverSheet.getDataRange().getValues();
        
        // หาไฟล์
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (row[0] === fileId) {
            const driveFileId = row[6];
            const driveFile = DriveApp.getFileById(driveFileId);
            
            return {
              success: true,
              data: {
                downloadUrl: `https://drive.google.com/uc?export=download&id=${driveFileId}`,
                filename: row[1]
              }
            };
          }
        }
      } catch (error) {
        console.error(`Error searching in ${server}:`, error);
        continue;
      }
    }
    
    return { success: false, message: 'ไม่พบไฟล์' };
  } catch (error) {
    console.error('Download error:', error);
    return { success: false, message: 'เกิดข้อผิดพลาดในการดาวน์โหลด: ' + error.toString() };
  }
}

// ฟังก์ชันลบไฟล์
function deleteFile(fileId, filename) {
  try {
    // ค้นหาและลบไฟล์ในทุก server tab
    for (const server of Object.keys(SERVER_FOLDERS)) {
      try {
        const serverSheet = getOrCreateSheet(SHEET_TABS[server]);
        const data = serverSheet.getDataRange().getValues();
        
        // หาและลบไฟล์
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (row[0] === fileId) {
            // ลบไฟล์จาก Drive
            try {
              const driveFile = DriveApp.getFileById(row[6]);
              driveFile.setTrashed(true);
            } catch (e) {
              console.log('File already deleted from Drive or not found');
            }
            
            // ลบแถวจาก Sheet
            serverSheet.deleteRow(i + 1);
            
            return { success: true, message: 'ลบไฟล์สำเร็จ' };
          }
        }
      } catch (error) {
        console.error(`Error deleting from ${server}:`, error);
        continue;
      }
    }
    
    return { success: false, message: 'ไม่พบไฟล์' };
  } catch (error) {
    console.error('Delete file error:', error);
    return { success: false, message: 'เกิดข้อผิดพลาดในการลบไฟล์: ' + error.toString() };
  }
}

// ฟังก์ชันโหลดรายการผู้ใช้
function getUsers() {
  try {
    const usersSheet = getOrCreateSheet(SHEET_TABS.users);
    const data = usersSheet.getDataRange().getValues();
    
    const users = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      users.push({
        username: row[0],
        role: row[2],
        status: row[3],
        created: formatDate(row[4])
      });
    }
    
    return { success: true, data: users };
  } catch (error) {
    console.error('Get users error:', error);
    return { success: false, message: 'เกิดข้อผิดพลาดในการโหลดผู้ใช้: ' + error.toString() };
  }
}

// ฟังก์ชันเพิ่มผู้ใช้
function addUser(username, password, role, status) {
  try {
    const usersSheet = getOrCreateSheet(SHEET_TABS.users);
    const data = usersSheet.getDataRange().getValues();
    
    // ตรวจสอบว่ามีผู้ใช้นี้แล้วหรือไม่
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        return { success: false, message: 'มีผู้ใช้นี้แล้ว' };
      }
    }
    
    // เพิ่มผู้ใช้ใหม่
    const newRow = [username, password, role, status, new Date()];
    usersSheet.appendRow(newRow);
    
    return { success: true, message: 'เพิ่มผู้ใช้สำเร็จ' };
  } catch (error) {
    console.error('Add user error:', error);
    return { success: false, message: 'เกิดข้อผิดพลาดในการเพิ่มผู้ใช้: ' + error.toString() };
  }
}

// ฟังก์ชันแก้ไขผู้ใช้
function updateUser(username, role, status) {
  try {
    const usersSheet = getOrCreateSheet(SHEET_TABS.users);
    const data = usersSheet.getDataRange().getValues();
    
    // หาและแก้ไขผู้ใช้
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        usersSheet.getRange(i + 1, 3).setValue(role);
        usersSheet.getRange(i + 1, 4).setValue(status);
        return { success: true, message: 'แก้ไขผู้ใช้สำเร็จ' };
      }
    }
    
    return { success: false, message: 'ไม่พบผู้ใช้' };
  } catch (error) {
    console.error('Update user error:', error);
    return { success: false, message: 'เกิดข้อผิดพลาดในการแก้ไขผู้ใช้: ' + error.toString() };
  }
}

// ฟังก์ชันรีเซ็ตรหัสผ่าน
function resetPassword(username, newPassword) {
  try {
    const usersSheet = getOrCreateSheet(SHEET_TABS.users);
    const data = usersSheet.getDataRange().getValues();
    
    // หาและแก้ไขรหัสผ่าน
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        usersSheet.getRange(i + 1, 2).setValue(newPassword);
        return { success: true, message: 'รีเซ็ตรหัสผ่านสำเร็จ' };
      }
    }
    
    return { success: false, message: 'ไม่พบผู้ใช้' };
  } catch (error) {
    console.error('Reset password error:', error);
    return { success: false, message: 'เกิดข้อผิดพลาดในการรีเซ็ตรหัสผ่าน: ' + error.toString() };
  }
}

// ฟังก์ชันลบผู้ใช้
function deleteUser(username) {
  try {
    const usersSheet = getOrCreateSheet(SHEET_TABS.users);
    const data = usersSheet.getDataRange().getValues();
    
    // หาและลบผู้ใช้
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        usersSheet.deleteRow(i + 1);
        return { success: true, message: 'ลบผู้ใช้สำเร็จ' };
      }
    }
    
    return { success: false, message: 'ไม่พบผู้ใช้' };
  } catch (error) {
    console.error('Delete user error:', error);
    return { success: false, message: 'เกิดข้อผิดพลาดในการลบผู้ใช้: ' + error.toString() };
  }
}

// ฟังก์ชันช่วยเหลือ
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function formatDate(date) {
  if (!date) return '';
  const d = new Date(date);
  return d.toLocaleDateString('th-TH') + ' ' + d.toLocaleTimeString('th-TH', {hour: '2-digit', minute: '2-digit'});
}