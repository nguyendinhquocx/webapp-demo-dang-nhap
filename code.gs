// code.gs
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl('https://gsheets.vn/wp-content/uploads/2024/05/cropped-42.png')
    .setTitle('Hệ thống quản lý nghỉ phép')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Thêm hàm mới để lấy danh sách người được phê duyệt
function getApproverRelationships(approverEmail) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User');
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var approvedEmails = [];
    
    // Bỏ qua hàng tiêu đề
    for (var i = 1; i < data.length; i++) {
      // Email người phê duyệt ở cột 8 (index 7)
      if (data[i][7] === approverEmail) {
        // Email của người được phê duyệt ở cột 3 (index 2)
        approvedEmails.push(data[i][2]);
      }
    }
    
    return approvedEmails;
  } catch (error) {
    console.error("Lỗi khi lấy danh sách người được phê duyệt:", error);
    return [];
  }
}

// Sửa lại hàm getTotalDataCounts để hỗ trợ việc phê duyệt
function getTotalDataCounts(email, isAdmin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pendingSheet = ss.getSheetByName('Đang xử lý');
  var approvedSheet = ss.getSheetByName('Phê duyệt');
  var disapprovedSheet = ss.getSheetByName('Huỷ bỏ');
  
  // Lấy dữ liệu từ các sheet
  var pendingData = pendingSheet.getDataRange().getValues();
  var approvedData = approvedSheet.getDataRange().getValues();
  var disapprovedData = disapprovedSheet.getDataRange().getValues();
  
  // Đếm số lượng theo email nếu không phải admin
  var pendingCount = 0;
  var approvedCount = 0;
  var disapprovedCount = 0;
  
  // Lấy thông tin người dùng
  var user = getUserByUsername(email);
  var isApprover = false;
  var approverFor = [];
  
  // Kiểm tra nếu người dùng là người phê duyệt
  if (user && user.role !== 'Admin') {
    var userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User');
    var userData = userSheet.getDataRange().getValues();
    
    for (var i = 1; i < userData.length; i++) {
      if (userData[i][7] === email) { // Email người phê duyệt ở cột 8 (index 7)
        isApprover = true;
        approverFor.push(userData[i][2]); // Email của người được phê duyệt
      }
    }
  }
  
  if (isAdmin === 'admin') {
    // Nếu là admin, đếm tất cả (trừ hàng tiêu đề)
    pendingCount = Math.max(0, pendingSheet.getLastRow() - 1);
    approvedCount = Math.max(0, approvedSheet.getLastRow() - 1);
    disapprovedCount = Math.max(0, disapprovedSheet.getLastRow() - 1);
  } else if (isApprover) {
    // Nếu là người phê duyệt, đếm cả các mục của họ và của người họ phê duyệt
    for (var i = 1; i < pendingData.length; i++) {
      if (pendingData[i][2] === email || approverFor.includes(pendingData[i][2])) pendingCount++;
    }
    
    for (var i = 1; i < approvedData.length; i++) {
      if (approvedData[i][2] === email || approverFor.includes(approvedData[i][2])) approvedCount++;
    }
    
    for (var i = 1; i < disapprovedData.length; i++) {
      if (disapprovedData[i][2] === email || approverFor.includes(disapprovedData[i][2])) disapprovedCount++;
    }
  } else {
    // Nếu là user thường, chỉ đếm các hàng có email trùng khớp
    // Email ở cột 3 (index 2)
    for (var i = 1; i < pendingData.length; i++) {
      if (pendingData[i][2] === email) pendingCount++;
    }
    
    for (var i = 1; i < approvedData.length; i++) {
      if (approvedData[i][2] === email) approvedCount++;
    }
    
    for (var i = 1; i < disapprovedData.length; i++) {
      if (disapprovedData[i][2] === email) disapprovedCount++;
    }
  }
  
  var data = {
    total: pendingCount + approvedCount + disapprovedCount,
    pending: pendingCount,
    approved: approvedCount,
    disapproved: disapprovedCount,
    approverFor: approverFor // Thêm danh sách người được phê duyệt
  };
  return data;
}

function checkUserHasData(email) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pendingSheet = ss.getSheetByName('Đang xử lý');
    var approvedSheet = ss.getSheetByName('Phê duyệt');
    var disapprovedSheet = ss.getSheetByName('Huỷ bỏ');
    
    // Kiểm tra từng sheet
    var pendingData = pendingSheet.getDataRange().getValues();
    for (var i = 1; i < pendingData.length; i++) {
      if (pendingData[i][2] === email) return true;
    }
    
    var approvedData = approvedSheet.getDataRange().getValues();
    for (var i = 1; i < approvedData.length; i++) {
      if (approvedData[i][2] === email) return true;
    }
    
    var disapprovedData = disapprovedSheet.getDataRange().getValues();
    for (var i = 1; i < disapprovedData.length; i++) {
      if (disapprovedData[i][2] === email) return true;
    }
    
    return false;
  } catch (error) {
    console.error("Lỗi kiểm tra dữ liệu người dùng:", error.message);
    return false; // Nếu có lỗi, trả về false để cho phép xóa
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function authenticate(username, password) {
  var userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User');
  var dataRange = userSheet.getDataRange();
  var values = dataRange.getValues();
  
  // Bỏ qua hàng tiêu đề (hàng đầu tiên)
  for (var i = 1; i < values.length; i++) {
    // Email ở cột 3 (index 2), Password ở cột 4 (index 3)
    var storedEmail = values[i][2];
    var storedPassword = values[i][3];
    var role = values[i][6]; // Role ở cột 7 (index 6) - Sau khi thêm Phòng ban và Email phê duyệt
    
    // Loại bỏ khoảng trắng và so sánh
    if (storedEmail && storedEmail.toString().trim() === username.trim() && 
        storedPassword && storedPassword.toString().trim() === password.trim()) {
      
      // Ghi log để debug (có thể gỡ bỏ sau)
      console.log("Đăng nhập thành công cho người dùng: " + username);
      
      if (role === "Admin") {
        return 'admin';
      } else {
        return 'user';
      }
    }
  }
  
  // Ghi log thất bại (có thể gỡ bỏ sau)
  console.log("Đăng nhập thất bại cho người dùng: " + username);
  return 'invalid';
}

function validateLogin(username, password) {
  // Đảm bảo input không có khoảng trắng thừa
  username = username.trim();
  password = password.trim();
  
  var validationResult = authenticate(username, password);
  return validationResult === 'user' || validationResult === 'admin' ? validationResult : 'invalid';
}

function getUserByUsername(username) {
  var sheetUser = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User');
  var dataUser = sheetUser.getDataRange().getValues();
  
  for (var i = 1; i < dataUser.length; i++) {
    if (dataUser[i][2] === username) { // Email đăng nhập ở cột thứ 3
      // Đảm bảo ngày tháng được định dạng chuẩn
      var leaveStartDate = dataUser[i][8];
      if (leaveStartDate instanceof Date) {
        leaveStartDate = Utilities.formatDate(leaveStartDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      
      return {
        id: dataUser[i][0],
        name: dataUser[i][1],
        email: dataUser[i][2],
        password: dataUser[i][3],
        image: dataUser[i][4],
        department: dataUser[i][5], // Phòng ban
        role: dataUser[i][6], // Chức vụ
        approverEmail: dataUser[i][7], // Email người phê duyệt
        leaveStartDate: leaveStartDate, // Đã định dạng chuẩn
        totalLeave: dataUser[i][9], // Tổng phép năm
        previousYearLeave: dataUser[i][10] || 0 // Phép năm trước chuyển sang
      };
    }
  }
  return null;
}

function checkDepartmentInUse(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = [
    ss.getSheetByName('Đang xử lý'),
    ss.getSheetByName('Phê duyệt'),
    ss.getSheetByName('Huỷ bỏ')
  ];
  
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    if (!sheet) continue;
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] === name) return true; // Phòng ban ở cột 4 (index 3)
    }
  }
  
  return false;
}

function checkRoleInUse(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = [
    ss.getSheetByName('Đang xử lý'),
    ss.getSheetByName('Phê duyệt'),
    ss.getSheetByName('Huỷ bỏ')
  ];
  
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    if (!sheet) continue;
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][4] === name) return true; // Chức vụ ở cột 5 (index 4)
    }
  }
  
  return false;
}

function getTotalCounts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pendingSheet = ss.getSheetByName('Đang xử lý');
  var approvedSheet = ss.getSheetByName('Phê duyệt');
  var disapprovedSheet = ss.getSheetByName('Huỷ bỏ');
  
  var pendingCount = Math.max(0, pendingSheet.getLastRow() - 1);
  var approvedCount = Math.max(0, approvedSheet.getLastRow() - 1);
  var disapprovedCount = Math.max(0, disapprovedSheet.getLastRow() - 1);
  
  return {
    pending: pendingCount,
    approved: approvedCount,
    disapproved: disapprovedCount
  };
}

function getTotalDataCounts(email, isAdmin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pendingSheet = ss.getSheetByName('Đang xử lý');
  var approvedSheet = ss.getSheetByName('Phê duyệt');
  var disapprovedSheet = ss.getSheetByName('Huỷ bỏ');
  
  // Lấy dữ liệu từ các sheet
  var pendingData = pendingSheet.getDataRange().getValues();
  var approvedData = approvedSheet.getDataRange().getValues();
  var disapprovedData = disapprovedSheet.getDataRange().getValues();
  
  // Đếm số lượng theo email nếu không phải admin
  var pendingCount = 0;
  var approvedCount = 0;
  var disapprovedCount = 0;
  
  // Lấy thông tin người dùng
  var user = getUserByUsername(email);
  var isApprover = false;
  var approverFor = [];
  
  // Kiểm tra nếu người dùng là người phê duyệt
  if (user && user.role !== 'Admin') {
    var userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User');
    var userData = userSheet.getDataRange().getValues();
    
    for (var i = 1; i < userData.length; i++) {
      if (userData[i][7] === email) { // Email người phê duyệt ở cột 8 (index 7)
        isApprover = true;
        approverFor.push(userData[i][2]); // Email của người được phê duyệt
      }
    }
  }
  
  if (isAdmin === 'admin') {
    // Nếu là admin, đếm tất cả (trừ hàng tiêu đề)
    pendingCount = Math.max(0, pendingSheet.getLastRow() - 1);
    approvedCount = Math.max(0, approvedSheet.getLastRow() - 1);
    disapprovedCount = Math.max(0, disapprovedSheet.getLastRow() - 1);
  } else if (isApprover) {
    // Nếu là người phê duyệt, đếm cả các mục của họ và của người họ phê duyệt
    for (var i = 1; i < pendingData.length; i++) {
      if (pendingData[i][2] === email || approverFor.includes(pendingData[i][2])) pendingCount++;
    }
    
    for (var i = 1; i < approvedData.length; i++) {
      if (approvedData[i][2] === email || approverFor.includes(approvedData[i][2])) approvedCount++;
    }
    
    for (var i = 1; i < disapprovedData.length; i++) {
      if (disapprovedData[i][2] === email || approverFor.includes(disapprovedData[i][2])) disapprovedCount++;
    }
  } else {
    // Nếu là user thường, chỉ đếm các hàng có email trùng khớp
    // Email ở cột 3 (index 2)
    for (var i = 1; i < pendingData.length; i++) {
      if (pendingData[i][2] === email) pendingCount++;
    }
    
    for (var i = 1; i < approvedData.length; i++) {
      if (approvedData[i][2] === email) approvedCount++;
    }
    
    for (var i = 1; i < disapprovedData.length; i++) {
      if (disapprovedData[i][2] === email) disapprovedCount++;
    }
  }
  
  var data = {
    total: pendingCount + approvedCount + disapprovedCount,
    pending: pendingCount,
    approved: approvedCount,
    disapproved: disapprovedCount
  };
  return data;
}

function getLeaveStatsByDepartment(email, userRole) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pendingSheet = ss.getSheetByName('Đang xử lý');
  var approvedSheet = ss.getSheetByName('Phê duyệt');
  var disapprovedSheet = ss.getSheetByName('Huỷ bỏ');
  
  var pendingData = pendingSheet.getDataRange().getValues();
  var approvedData = approvedSheet.getDataRange().getValues();
  var disapprovedData = disapprovedSheet.getDataRange().getValues();
  
  // Lấy danh sách phòng ban
  var departments = getDepartments();
  
  // Lấy thông tin người dùng
  var user = getUserByUsername(email);
  var isApprover = false;
  var approverFor = [];
  
  // Tạo đối tượng lưu trữ thống kê theo phòng ban
  var statsByDepartment = {};
  
  // Kiểm tra xem người dùng có phải là người phê duyệt không
  if (user && user.role !== 'Admin') {
    var userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User');
    var userData = userSheet.getDataRange().getValues();
    
    for (var i = 1; i < userData.length; i++) {
      if (userData[i][7] === email) { // Email người phê duyệt ở cột 8 (index 7)
        isApprover = true;
        approverFor.push(userData[i][2]); // Email của người được phê duyệt
      }
    }
  }
  
  // Nếu là admin, hiển thị tất cả phòng ban
  if (userRole === 'admin') {
    departments.forEach(function(dept) {
      statsByDepartment[dept] = {
        pending: 0,
        approved: 0,
        disapproved: 0
      };
    });
  } else if (isApprover) {
    // Nếu là người phê duyệt, hiển thị phòng ban của người dùng và những người họ phê duyệt
    statsByDepartment[user.department] = {
      pending: 0,
      approved: 0,
      disapproved: 0
    };
    
    // Thêm các phòng ban của người được phê duyệt
    for (var i = 1; i < userData.length; i++) {
      if (userData[i][7] === email) { // Email phê duyệt
        var empDepartment = userData[i][5]; // Phòng ban ở cột 6 (index 5)
        if (!statsByDepartment[empDepartment]) {
          statsByDepartment[empDepartment] = {
            pending: 0,
            approved: 0,
            disapproved: 0
          };
        }
      }
    }
  } else {
    // Nếu là nhân viên thường, chỉ hiển thị thống kê cá nhân, không phải toàn phòng ban
    // Tạo một key đặc biệt cho người dùng đó, ví dụ: "Cá nhân"
    statsByDepartment["Cá nhân"] = {
      pending: 0,
      approved: 0,
      disapproved: 0
    };
  }
  
  // Đếm đơn đang xử lý
  for (var i = 1; i < pendingData.length; i++) {
    var dept = pendingData[i][3]; // Phòng ban ở cột 4 (index 3)
    var requestEmail = pendingData[i][2]; // Email người yêu cầu
    var leaveDays = calculateLeaveDays(pendingData[i][5], pendingData[i][6], pendingData[i][7], pendingData[i][8]);
    
    if (userRole === 'admin') {
      // Admin thấy tất cả theo phòng ban
      if (statsByDepartment[dept] !== undefined) {
        statsByDepartment[dept].pending += leaveDays;
      } else {
        statsByDepartment[dept] = {
          pending: leaveDays,
          approved: 0,
          disapproved: 0
        };
      }
    } else if (isApprover) {
      // Người phê duyệt thấy dữ liệu của mình và người họ phê duyệt
      if (requestEmail === email) {
        // Dữ liệu của chính họ
        if (statsByDepartment[dept] !== undefined) {
          statsByDepartment[dept].pending += leaveDays;
        }
      } else if (approverFor.includes(requestEmail)) {
        // Dữ liệu của người họ phê duyệt
        if (statsByDepartment[dept] !== undefined) {
          statsByDepartment[dept].pending += leaveDays;
        }
      }
    } else {
      // User thường chỉ thấy dữ liệu của mình
      if (requestEmail === email) {
        statsByDepartment["Cá nhân"].pending += leaveDays;
      }
    }
  }
  
  // Đếm đơn đã phê duyệt
  for (var i = 1; i < approvedData.length; i++) {
    var dept = approvedData[i][3];
    var requestEmail = approvedData[i][2];
    var leaveDays = calculateLeaveDays(approvedData[i][5], approvedData[i][6], approvedData[i][7], approvedData[i][8]);
    
    if (userRole === 'admin') {
      // Admin thấy tất cả theo phòng ban
      if (statsByDepartment[dept] !== undefined) {
        statsByDepartment[dept].approved += leaveDays;
      } else {
        statsByDepartment[dept] = {
          pending: 0,
          approved: leaveDays,
          disapproved: 0
        };
      }
    } else if (isApprover) {
      // Người phê duyệt thấy dữ liệu của mình và người họ phê duyệt
      if (requestEmail === email) {
        // Dữ liệu của chính họ
        if (statsByDepartment[dept] !== undefined) {
          statsByDepartment[dept].approved += leaveDays;
        }
      } else if (approverFor.includes(requestEmail)) {
        // Dữ liệu của người họ phê duyệt
        if (statsByDepartment[dept] !== undefined) {
          statsByDepartment[dept].approved += leaveDays;
        }
      }
    } else {
      // User thường chỉ thấy dữ liệu của mình
      if (requestEmail === email) {
        statsByDepartment["Cá nhân"].approved += leaveDays;
      }
    }
  }
  
  // Đếm đơn đã huỷ bỏ
  for (var i = 1; i < disapprovedData.length; i++) {
    var dept = disapprovedData[i][3];
    var requestEmail = disapprovedData[i][2];
    var leaveDays = calculateLeaveDays(disapprovedData[i][5], disapprovedData[i][6], disapprovedData[i][7], disapprovedData[i][8]);
    
    if (userRole === 'admin') {
      // Admin thấy tất cả theo phòng ban
      if (statsByDepartment[dept] !== undefined) {
        statsByDepartment[dept].disapproved += leaveDays;
      } else {
        statsByDepartment[dept] = {
          pending: 0,
          approved: 0,
          disapproved: leaveDays
        };
      }
    } else if (isApprover) {
      // Người phê duyệt thấy dữ liệu của mình và người họ phê duyệt
      if (requestEmail === email) {
        // Dữ liệu của chính họ
        if (statsByDepartment[dept] !== undefined) {
          statsByDepartment[dept].disapproved += leaveDays;
        }
      } else if (approverFor.includes(requestEmail)) {
        // Dữ liệu của người họ phê duyệt
        if (statsByDepartment[dept] !== undefined) {
          statsByDepartment[dept].disapproved += leaveDays;
        }
      }
    } else {
      // User thường chỉ thấy dữ liệu của mình
      if (requestEmail === email) {
        statsByDepartment["Cá nhân"].disapproved += leaveDays;
      }
    }
  }
  
  // Chuyển đổi sang mảng để dễ sử dụng
  var result = [];
  for (var dept in statsByDepartment) {
    // Chỉ thêm các phòng ban có dữ liệu
    if (statsByDepartment[dept].pending > 0 || 
        statsByDepartment[dept].approved > 0 || 
        statsByDepartment[dept].disapproved > 0) {
      result.push({
        department: dept,
        pending: statsByDepartment[dept].pending,
        approved: statsByDepartment[dept].approved,
        disapproved: statsByDepartment[dept].disapproved,
        total: statsByDepartment[dept].pending + statsByDepartment[dept].approved + statsByDepartment[dept].disapproved
      });
    }
  }
  
  return result;
}

function getLeaveStatsByMonth(email, userRole) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pendingSheet = ss.getSheetByName('Đang xử lý');
  var approvedSheet = ss.getSheetByName('Phê duyệt');
  var disapprovedSheet = ss.getSheetByName('Huỷ bỏ');
  
  var pendingData = pendingSheet.getDataRange().getValues();
  var approvedData = approvedSheet.getDataRange().getValues();
  var disapprovedData = disapprovedSheet.getDataRange().getValues();
  
  // Lấy thông tin người dùng
  var user = getUserByUsername(email);
  var isApprover = false;
  var approverFor = [];
  
  // Kiểm tra nếu người dùng là người phê duyệt
  if (user && user.role !== 'Admin') {
    var userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User');
    var userData = userSheet.getDataRange().getValues();
    
    for (var i = 1; i < userData.length; i++) {
      if (userData[i][7] === email) { // Email người phê duyệt ở cột 8 (index 7)
        isApprover = true;
        approverFor.push(userData[i][2]); // Email của người được phê duyệt
      }
    }
  }
  
  // Tạo đối tượng lưu trữ thống kê theo tháng
  var statsByMonth = {};
  var months = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'];
  var currentYear = new Date().getFullYear();
  
  // Khởi tạo với tất cả các tháng trong năm hiện tại
  months.forEach(function(month) {
    statsByMonth[month + '/' + currentYear] = {
      pending: 0,
      approved: 0,
      disapproved: 0
    };
  });
  
  // Hàm trích xuất tháng/năm từ ngày
  function getMonthYear(dateStr) {
    var date = new Date(dateStr);
    var month = ('0' + (date.getMonth() + 1)).slice(-2);
    var year = date.getFullYear();
    return month + '/' + year;
  }
  
  // Xử lý dữ liệu đang xử lý
  for (var i = 1; i < pendingData.length; i++) {
    try {
      var startDate = pendingData[i][5]; // Ngày bắt đầu ở cột 6 (index 5)
      var dept = pendingData[i][3]; // Phòng ban
      var requestEmail = pendingData[i][2]; // Email người yêu cầu
      var leaveDays = calculateLeaveDays(pendingData[i][5], pendingData[i][6], pendingData[i][7], pendingData[i][8]);
      
      // Chuyển đổi định dạng ngày nếu cần
      if (typeof startDate === 'string' && startDate.includes('/')) {
        // Nếu ngày ở định dạng dd/MM/yyyy
        var parts = startDate.split('/');
        startDate = new Date(parts[2], parts[1] - 1, parts[0]);
      } else if (startDate instanceof Date) {
        // Giữ nguyên nếu đã là Date
      } else {
        // Thử chuyển đổi
        startDate = new Date(startDate);
      }
      
      var monthYear = getMonthYear(startDate);
      
      if (userRole === 'admin') {
        // Admin thấy tất cả dữ liệu
        if (statsByMonth[monthYear]) {
          statsByMonth[monthYear].pending += leaveDays;
        } else {
          statsByMonth[monthYear] = {
            pending: leaveDays,
            approved: 0,
            disapproved: 0
          };
        }
      } else if (isApprover) {
        // Người phê duyệt thấy dữ liệu của mình và người họ phê duyệt
        if (requestEmail === email || approverFor.includes(requestEmail)) {
          if (statsByMonth[monthYear]) {
            statsByMonth[monthYear].pending += leaveDays;
          } else {
            statsByMonth[monthYear] = {
              pending: leaveDays,
              approved: 0,
              disapproved: 0
            };
          }
        }
      } else {
        // User thường chỉ thấy dữ liệu của mình
        if (requestEmail === email) {
          if (statsByMonth[monthYear]) {
            statsByMonth[monthYear].pending += leaveDays;
          } else {
            statsByMonth[monthYear] = {
              pending: leaveDays,
              approved: 0,
              disapproved: 0
            };
          }
        }
      }
    } catch (error) {
      console.error("Lỗi khi xử lý dữ liệu đang xử lý:", error, "Dòng:", i);
    }
  }
  
  // Xử lý dữ liệu đã phê duyệt
  for (var i = 1; i < approvedData.length; i++) {
    try {
      var startDate = approvedData[i][5];
      var dept = approvedData[i][3];
      var requestEmail = approvedData[i][2];
      var leaveDays = calculateLeaveDays(approvedData[i][5], approvedData[i][6], approvedData[i][7], approvedData[i][8]);
      
      // Chuyển đổi định dạng ngày nếu cần
      if (typeof startDate === 'string' && startDate.includes('/')) {
        // Nếu ngày ở định dạng dd/MM/yyyy
        var parts = startDate.split('/');
        startDate = new Date(parts[2], parts[1] - 1, parts[0]);
      } else if (startDate instanceof Date) {
        // Giữ nguyên nếu đã là Date
      } else {
        // Thử chuyển đổi
        startDate = new Date(startDate);
      }
      
      var monthYear = getMonthYear(startDate);
      
      if (userRole === 'admin') {
        // Admin thấy tất cả dữ liệu
        if (statsByMonth[monthYear]) {
          statsByMonth[monthYear].approved += leaveDays;
        } else {
          statsByMonth[monthYear] = {
            pending: 0,
            approved: leaveDays,
            disapproved: 0
          };
        }
      } else if (isApprover) {
        // Người phê duyệt thấy dữ liệu của mình và người họ phê duyệt
        if (requestEmail === email || approverFor.includes(requestEmail)) {
          if (statsByMonth[monthYear]) {
            statsByMonth[monthYear].approved += leaveDays;
          } else {
            statsByMonth[monthYear] = {
              pending: 0,
              approved: leaveDays,
              disapproved: 0
            };
          }
        }
      } else {
        // User thường chỉ thấy dữ liệu của mình
        if (requestEmail === email) {
          if (statsByMonth[monthYear]) {
            statsByMonth[monthYear].approved += leaveDays;
          } else {
            statsByMonth[monthYear] = {
              pending: 0,
              approved: leaveDays,
              disapproved: 0
            };
          }
        }
      }
    } catch (error) {
      console.error("Lỗi khi xử lý dữ liệu đã phê duyệt:", error, "Dòng:", i);
    }
  }
  
  // Xử lý dữ liệu đã huỷ bỏ
  for (var i = 1; i < disapprovedData.length; i++) {
    try {
      var startDate = disapprovedData[i][5];
      var dept = disapprovedData[i][3];
      var requestEmail = disapprovedData[i][2];
      var leaveDays = calculateLeaveDays(disapprovedData[i][5], disapprovedData[i][6], disapprovedData[i][7], disapprovedData[i][8]);
      
      // Chuyển đổi định dạng ngày nếu cần
      if (typeof startDate === 'string' && startDate.includes('/')) {
        // Nếu ngày ở định dạng dd/MM/yyyy
        var parts = startDate.split('/');
        startDate = new Date(parts[2], parts[1] - 1, parts[0]);
      } else if (startDate instanceof Date) {
        // Giữ nguyên nếu đã là Date
      } else {
        // Thử chuyển đổi
        startDate = new Date(startDate);
      }
      
      var monthYear = getMonthYear(startDate);
      
      if (userRole === 'admin') {
        // Admin thấy tất cả dữ liệu
        if (statsByMonth[monthYear]) {
          statsByMonth[monthYear].disapproved += leaveDays;
        } else {
          statsByMonth[monthYear] = {
            pending: 0,
            approved: 0,
            disapproved: leaveDays
          };
        }
      } else if (isApprover) {
        // Người phê duyệt thấy dữ liệu của mình và người họ phê duyệt
        if (requestEmail === email || approverFor.includes(requestEmail)) {
          if (statsByMonth[monthYear]) {
            statsByMonth[monthYear].disapproved += leaveDays;
          } else {
            statsByMonth[monthYear] = {
              pending: 0,
              approved: 0,
              disapproved: leaveDays
            };
          }
        }
      } else {
        // User thường chỉ thấy dữ liệu của mình
        if (requestEmail === email) {
          if (statsByMonth[monthYear]) {
            statsByMonth[monthYear].disapproved += leaveDays;
          } else {
            statsByMonth[monthYear] = {
              pending: 0,
              approved: 0,
              disapproved: leaveDays
            };
          }
        }
      }
    } catch (error) {
      console.error("Lỗi khi xử lý dữ liệu đã huỷ bỏ:", error, "Dòng:", i);
    }
  }
  
  // Chuyển đổi sang mảng để dễ sử dụng
  var result = [];
  for (var monthYear in statsByMonth) {
    // Chỉ thêm các tháng có dữ liệu
    if (statsByMonth[monthYear].pending > 0 || 
        statsByMonth[monthYear].approved > 0 || 
        statsByMonth[monthYear].disapproved > 0) {
      var parts = monthYear.split('/');
      var monthName = getMonthName(parseInt(parts[0]));
      
      result.push({
        month: monthName,
        value: parts[0], // Giữ lại số tháng để sắp xếp
        pending: statsByMonth[monthYear].pending,
        approved: statsByMonth[monthYear].approved,
        disapproved: statsByMonth[monthYear].disapproved,
        total: statsByMonth[monthYear].pending + statsByMonth[monthYear].approved + statsByMonth[monthYear].disapproved
      });
    }
  }
  
  // Sắp xếp theo thứ tự tháng
  result.sort(function(a, b) {
    return parseInt(a.value) - parseInt(b.value);
  });
  
  return result;
}

// Hàm hỗ trợ lấy tên tháng
function getMonthName(monthNumber) {
  var months = ['Tháng 1', 'Tháng 2', 'Tháng 3', 'Tháng 4', 'Tháng 5', 'Tháng 6', 
                'Tháng 7', 'Tháng 8', 'Tháng 9', 'Tháng 10', 'Tháng 11', 'Tháng 12'];
  return months[monthNumber - 1];
}

// Tính toán số ngày nghỉ phép
function calculateLeaveDays(startDate, endDate, leaveType, leaveSession) {
  if (leaveType === "Trong ngày") {
    if (leaveSession === "Cả ngày") {
      // Kiểm tra nếu ngày này là ngày lễ
      if (isHoliday(startDate)) {
        return 0.0; // Không tính ngày phép nếu là ngày lễ
      }
      return 1.0;
    } else if (leaveSession === "Buổi sáng" || leaveSession === "Buổi chiều") {
      // Kiểm tra nếu ngày này là ngày lễ
      if (isHoliday(startDate)) {
        return 0.0; // Không tính ngày phép nếu là ngày lễ
      }
      return 0.5;
    }
  } else if (leaveType === "Từ ngày đến ngày") {
    var start = new Date(startDate);
    var end = new Date(endDate);
    var dayMilliseconds = 1000 * 60 * 60 * 24;
    
    // Cộng thêm 1 ngày vì tính cả ngày đầu và cuối
    var diffDays = Math.round((end - start) / dayMilliseconds) + 1;
    
    // Trừ đi các ngày cuối tuần (thứ 7, chủ nhật)
    for (var day = new Date(start); day <= end; day.setDate(day.getDate() + 1)) {
      var dayOfWeek = day.getDay(); // 0 = Chủ nhật, 6 = Thứ 7
      if (dayOfWeek === 0 || dayOfWeek === 6) {
        diffDays--;
      }
    }
    
    // Kiểm tra trùng với ngày nghỉ lễ
    var holidays = getHolidays();
    for (var i = 0; i < holidays.length; i++) {
      // Chuyển đổi định dạng ngày lễ (dd/MM/yyyy) thành đối tượng Date
      var holidayParts = holidays[i].split('/');
      if (holidayParts.length !== 3) continue;
      
      // Chú ý: JS tháng bắt đầu từ 0, nên phải trừ 1
      var holiday = new Date(
        parseInt(holidayParts[2]), // năm
        parseInt(holidayParts[1]) - 1, // tháng (0-11)
        parseInt(holidayParts[0]) // ngày
      );
      
      // So sánh với khoảng thời gian
      if (holiday >= start && holiday <= end && 
          holiday.getDay() !== 0 && holiday.getDay() !== 6) { // Chỉ trừ nếu ngày lễ không rơi vào cuối tuần
        diffDays--;
        console.log("Trừ 1 ngày nghỉ lễ:", holidays[i]);
      }
    }
    
    return Math.max(0, diffDays);
  }
  
  return 0; // Mặc định nếu không xác định được
}

function isHoliday(dateValue) {
  var checkDate;
  
  // Chuyển đổi input date thành đối tượng Date
  if (dateValue instanceof Date) {
    checkDate = new Date(dateValue);
  } else if (typeof dateValue === 'string') {
    // Xử lý chuỗi định dạng yyyy-MM-dd (từ input type="date")
    if (dateValue.includes('-')) {
      var parts = dateValue.split('-');
      checkDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
    } 
    // Xử lý chuỗi định dạng dd/MM/yyyy (định dạng hiển thị)
    else if (dateValue.includes('/')) {
      var parts = dateValue.split('/');
      checkDate = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
    }
    else {
      checkDate = new Date(dateValue);
    }
  } else {
    return false; // Không thể xác định định dạng ngày
  }
  
  // Đảm bảo giờ là 00:00:00 để so sánh chính xác
  checkDate.setHours(0, 0, 0, 0);
  
  // Lấy danh sách ngày lễ
  var holidays = getHolidays();
  
  // Kiểm tra từng ngày lễ
  for (var i = 0; i < holidays.length; i++) {
    var holidayStr = holidays[i];
    var holidayParts = holidayStr.split('/');
    
    if (holidayParts.length !== 3) continue;
    
    // Tạo đối tượng Date từ chuỗi dd/MM/yyyy
    var holidayDate = new Date(
      parseInt(holidayParts[2]), // năm
      parseInt(holidayParts[1]) - 1, // tháng (0-11)
      parseInt(holidayParts[0]) // ngày
    );
    
    // So sánh ngày (bỏ qua giờ, phút, giây)
    if (holidayDate.getDate() === checkDate.getDate() && 
        holidayDate.getMonth() === checkDate.getMonth() && 
        holidayDate.getFullYear() === checkDate.getFullYear()) {
      return true;
    }
  }
  
  return false;
}

// Hàm lấy danh sách ngày nghỉ lễ
function getHolidays() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) return [];
  
  var lastRow = Math.max(sheet.getLastRow(), 1);
  // Bỏ qua dòng tiêu đề
  if (lastRow <= 1) return [];
  
  var data = sheet.getRange('C2:C' + lastRow).getValues();
  var holidays = [];
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] !== "") {
      // Nếu dữ liệu là Date, định dạng thành chuỗi ngày tháng
      if (data[i][0] instanceof Date) {
        var dateStr = Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), "dd/MM/yyyy");
        holidays.push(dateStr);
      } else {
        holidays.push(data[i][0]);
      }
    }
  }
  
  Logger.log("Holidays retrieved: " + holidays.length); // Thêm log để debug
  return holidays;
}

// Tính số ngày phép đã sử dụng và còn lại
function calculateLeaveBalance(email) {
  var user = getUserByUsername(email);
  if (!user) return { 
    used: 0, 
    remaining: 0, 
    total: 0, 
    previousYear: 0, 
    previousYearExpiry: getPreviousYearLeaveExpiryDate(),
    remainingCurrentYear: 0,
    usedPreviousYear: 0,
    canUsePreviousYear: false
  };
  
  var totalLeave = parseFloat(user.totalLeave) || 0;
  var previousYearLeave = parseFloat(user.previousYearLeave) || 0;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pendingSheet = ss.getSheetByName('Đang xử lý');
  var approvedSheet = ss.getSheetByName('Phê duyệt');
  
  var pendingData = pendingSheet.getDataRange().getValues();
  var approvedData = approvedSheet.getDataRange().getValues();
  
  var usedLeave = 0;
  var usedPreviousYearLeave = 0;
  
  // Lấy ngày hiện tại
  var today = new Date();
  
  // Lấy ngày hết hạn từ cài đặt
  var expiryDateStr = getPreviousYearLeaveExpiryDate();
  var expireParts = expiryDateStr.split('/');
  var expireDate = new Date(
    parseInt(expireParts[2]), // năm
    parseInt(expireParts[1]) - 1, // tháng (0-11)
    parseInt(expireParts[0]), // ngày
    23, 59, 59 // giờ, phút, giây
  );
  
  // Kiểm tra nếu ngày hiện tại vượt quá ngày hết hạn
  var canUsePreviousYear = today <= expireDate;
  
  // Flag để phân biệt giữa các đơn của năm trước và năm nay
  var previousYearRequests = [];
  var currentYearRequests = [];
  
  // Hàm phân loại đơn theo thời gian
  function categorizeLeaveRequest(rowData) {
    if (rowData[2] === email) { // Email ở cột 3 (index 2)
      var startDate = parseDate(rowData[5]); // Ngày bắt đầu
      
      // Thêm vào danh sách tương ứng
      if (startDate <= expireDate) {
        previousYearRequests.push(rowData);
      } else {
        currentYearRequests.push(rowData);
      }
    }
  }
  
  // Phân loại đơn đang xử lý
  for (var i = 1; i < pendingData.length; i++) {
    categorizeLeaveRequest(pendingData[i]);
  }
  
  // Phân loại đơn đã phê duyệt
  for (var i = 1; i < approvedData.length; i++) {
    categorizeLeaveRequest(approvedData[i]);
  }
  
  // Xử lý đơn năm trước (trước hoặc vào ngày hết hạn) - ưu tiên dùng phép năm trước
  for (var i = 0; i < previousYearRequests.length; i++) {
    var rowData = previousYearRequests[i];
    var leaveType = rowData[7]; // Loại nghỉ ở cột 8 (index 7)
    var leaveSession = rowData[8]; // Buổi nghỉ ở cột 9 (index 8)
    var leaveDays = calculateLeaveDays(rowData[5], rowData[6], leaveType, leaveSession);
    
    // Ưu tiên sử dụng phép năm trước
    var availablePreviousLeave = Math.max(0, previousYearLeave - usedPreviousYearLeave);
    
    if (availablePreviousLeave >= leaveDays) {
      usedPreviousYearLeave += leaveDays;
    } else {
      // Nếu phép năm trước không đủ, dùng hết phép năm trước rồi dùng phép năm nay
      usedPreviousYearLeave += availablePreviousLeave;
      usedLeave += (leaveDays - availablePreviousLeave);
    }
  }
  
  // Xử lý đơn năm nay (sau ngày hết hạn) - chỉ dùng phép năm nay
  for (var i = 0; i < currentYearRequests.length; i++) {
    var rowData = currentYearRequests[i];
    var leaveType = rowData[7]; // Loại nghỉ ở cột 8 (index 7)
    var leaveSession = rowData[8]; // Buổi nghỉ ở cột 9 (index 8)
    var leaveDays = calculateLeaveDays(rowData[5], rowData[6], leaveType, leaveSession);
    
    // Chỉ sử dụng phép năm nay
    usedLeave += leaveDays;
  }
  
  // Tính số phép còn lại
  var remainingPreviousLeave = Math.max(0, previousYearLeave - usedPreviousYearLeave);
  var remainingCurrentLeave = Math.max(0, totalLeave - usedLeave);
  
  // Tổng phép còn lại (nếu còn trong thời hạn sử dụng phép năm trước thì cộng thêm)
  var remainingLeave = remainingCurrentLeave + (canUsePreviousYear ? remainingPreviousLeave : 0);
  
  return {
    used: usedLeave,
    usedPreviousYear: usedPreviousYearLeave,
    remaining: remainingLeave,
    remainingPreviousYear: remainingPreviousLeave,
    remainingCurrentYear: remainingCurrentLeave,
    total: totalLeave,
    previousYear: previousYearLeave,
    previousYearExpiry: expiryDateStr,
    canUsePreviousYear: canUsePreviousYear
  };
}

function parseDate(dateStr) {
  // Nếu đã là đối tượng Date
  if (dateStr instanceof Date) {
    return dateStr;
  }
  
  // Nếu là chuỗi định dạng dd/MM/yyyy
  if (typeof dateStr === 'string' && dateStr.includes('/')) {
    var parts = dateStr.split('/');
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  
  // Các trường hợp khác
  return new Date(dateStr);
}

function addLeaveRequest(form) {
  try {
    var folderName = '📁Lưu tệp V3';
    var folder;
    var folderIterator = DriveApp.getFoldersByName(folderName);
    if (folderIterator.hasNext()) {
      folder = folderIterator.next();
    } else {
      folder = DriveApp.createFolder(folderName);
      Logger.log('Đã tạo thư mục mới: ' + folderName);
    }
    
    // Các thông tin cơ bản
    var name = form.name;
    var email = form.email;
    var department = form.department;
    var role = form.role;
    var startDate, endDate, leaveSession, leaveType;
    
    // Xác định loại nghỉ phép và thời gian
    leaveType = form.leaveType;
    
    if (leaveType === 'Trong ngày') {
      startDate = Utilities.formatDate(new Date(form.startDate), Session.getScriptTimeZone(), "dd/MM/yyyy");
      endDate = startDate;
      leaveSession = form.leaveSession || "Cả ngày";
    } else if (leaveType === 'Từ ngày đến ngày') {
      startDate = Utilities.formatDate(new Date(form.startDateRange), Session.getScriptTimeZone(), "dd/MM/yyyy");
      endDate = Utilities.formatDate(new Date(form.endDate), Session.getScriptTimeZone(), "dd/MM/yyyy");
      leaveSession = "Cả ngày";
    } else {
      throw new Error("Loại nghỉ phép không hợp lệ");
    }
    
    var reason = form.reason;
    var status = form.status || "Đang xử lý";
    var note = form.note || "-";
    
    // Kiểm tra trùng lặp các ngày nghỉ
    var hasOverlap = checkLeaveOverlap(email, 
      leaveType === 'Trong ngày' ? form.startDate : form.startDateRange,
      leaveType === 'Trong ngày' ? form.startDate : form.endDate,
      leaveType, leaveSession);
    
    if (hasOverlap) {
      throw new Error("Bạn đã có đơn nghỉ phép cho thời gian này. Vui lòng kiểm tra lại!");
    }
    
    // Chuyển đổi chuỗi ngày thành đối tượng Date để so sánh
    var startDateObj = new Date(leaveType === 'Trong ngày' ? form.startDate : form.startDateRange);
    var endDateObj = new Date(leaveType === 'Trong ngày' ? form.startDate : form.endDate);
    
    // Lấy ngày hết hạn phép năm trước từ cài đặt
    var expiryDateStr = getPreviousYearLeaveExpiryDate();
    var expireParts = expiryDateStr.split('/');
    var expireYear = parseInt(expireParts[2]);
    var expireMonth = parseInt(expireParts[1]) - 1; // Chuyển về index 0-11
    var expireDay = parseInt(expireParts[0]);
    var expireDate = new Date(expireYear, expireMonth, expireDay, 23, 59, 59);
    
    // Tính số ngày nghỉ
    var leaveDays = calculateLeaveDays(
      leaveType === 'Trong ngày' ? form.startDate : form.startDateRange, 
      leaveType === 'Trong ngày' ? form.startDate : form.endDate, 
      leaveType, 
      leaveSession
    );
    
    // Kiểm tra số phép còn lại
    var leaveBalance = calculateLeaveBalance(email);
    
    // Xử lý đặc biệt cho trường hợp khoảng thời gian chuyển tiếp (từ trước hết hạn đến sau hết hạn)
    if (leaveType === 'Từ ngày đến ngày' && startDateObj <= expireDate && endDateObj > expireDate) {
      // Đây là trường hợp khoảng thời gian chuyển tiếp
      // Tính số ngày trước hoặc vào ngày hết hạn
      var tempEndDate = new Date(expireYear, expireMonth, expireDay); // Ngày hết hạn
      var daysBeforeExpire = calculateLeaveDays(startDateObj, tempEndDate, "Từ ngày đến ngày", "Cả ngày");
      
      // Tính số ngày sau ngày hết hạn
      var tempStartDate = new Date(expireYear, expireMonth, expireDay);
      tempStartDate.setDate(tempStartDate.getDate() + 1); // Ngày sau ngày hết hạn
      var daysAfterExpire = calculateLeaveDays(tempStartDate, endDateObj, "Từ ngày đến ngày", "Cả ngày");
      
      // Kiểm tra phép năm trước (đủ cho phần trước ngày hết hạn không)
      if (daysBeforeExpire > leaveBalance.remainingPreviousYear) {
        throw new Error("Số ngày nghỉ trước " + expiryDateStr + " (" + daysBeforeExpire + 
                       " ngày) vượt quá số phép năm trước còn lại (" + 
                       leaveBalance.remainingPreviousYear + " ngày).");
      }
      
      // Kiểm tra phép năm nay (đủ cho phần sau ngày hết hạn không)
      if (daysAfterExpire > leaveBalance.remainingCurrentYear) {
        throw new Error("Số ngày nghỉ sau " + expiryDateStr + " (" + daysAfterExpire + 
                       " ngày) vượt quá số phép năm nay còn lại (" + 
                       leaveBalance.remainingCurrentYear + " ngày).");
      }
      
      // Nếu cả hai điều kiện đều thỏa mãn, chúng ta tiếp tục
      console.log("Trường hợp chuyển tiếp: " + daysBeforeExpire + " ngày trước " + expiryDateStr + " và " + 
                  daysAfterExpire + " ngày sau " + expiryDateStr);
    } else {
      // Trường hợp bình thường (không nằm ở khoảng thời gian chuyển tiếp)
      // Xác định số phép có thể sử dụng dựa vào ngày bắt đầu nghỉ
      var availableLeave;
      
      // Đảm bảo so sánh chính xác với ngày hết hạn (làm tròn xuống 00:00:00)
      var compareStartDate = new Date(startDateObj.getFullYear(), startDateObj.getMonth(), startDateObj.getDate());
      var compareExpireDate = new Date(expireDate.getFullYear(), expireDate.getMonth(), expireDate.getDate());
      
      if (compareStartDate > compareExpireDate) {
        // Nếu nghỉ sau ngày hết hạn, chỉ tính phép năm nay
        availableLeave = leaveBalance.remainingCurrentYear;
        
        if (leaveDays > availableLeave) {
          throw new Error("Số ngày nghỉ (" + leaveDays + 
                        " ngày) vượt quá số phép năm nay còn lại (" + 
                        availableLeave + " ngày).");
        }
      } else {
        // Nếu nghỉ trước hoặc vào ngày hết hạn, ưu tiên dùng phép năm trước
        if (leaveBalance.remainingPreviousYear >= leaveDays) {
          // Nếu phép năm trước đủ, dùng phép năm trước
          // Không cần kiểm tra thêm
        } else {
          // Nếu phép năm trước không đủ, kiểm tra tổng phép
          if (leaveDays > leaveBalance.remaining) {
            throw new Error("Số ngày nghỉ (" + leaveDays + 
                          " ngày) vượt quá tổng số phép còn lại (" + 
                          leaveBalance.remaining + " ngày).");
          }
        }
      }
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Đang xử lý');
    var lastRow = sheet.getLastRow();
    var id;
    
    function generateID() {
      var randomId = '1' + Math.floor(10000 + Math.random() * 90000);
      var existingIds = sheet.getRange("A2:A" + lastRow).getValues().flat();
      while (existingIds.includes("'" + randomId)) {
        randomId = '1' + Math.floor(10000 + Math.random() * 90000);
      }
      return "'" + randomId;
    }
    
    id = generateID();
    var fileUrl = "";
    
    // Xử lý tệp đính kèm từ form
    var fileBlob = form.myFile;
    
    if (fileBlob && fileBlob.getName && fileBlob.getName()) {
      // Chấp nhận PDF, hình ảnh, Word và Excel
      if (fileBlob.getContentType().startsWith('application/pdf') || 
          fileBlob.getContentType().startsWith('image/') ||
          fileBlob.getContentType().includes('word') ||
          fileBlob.getContentType().includes('excel') ||
          fileBlob.getContentType().includes('spreadsheet')) {
        
        try {
          // Lưu file vào Drive
          var file = folder.createFile(fileBlob);
          
          // Cài đặt quyền truy cập "Anyone with the link can view"
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          
          // Lưu URL truy cập
          fileUrl = file.getUrl();
          
          console.log("Đã tạo file: " + fileUrl);
        } catch (fileError) {
          console.error("Lỗi khi tạo file: " + fileError.toString());
          // Vẫn tiếp tục thêm dòng, chỉ ghi nhận lỗi
        }
      } else {
        throw new Error("Loại tệp không hợp lệ. Chỉ chấp nhận PDF, hình ảnh, Word và Excel.");
      }
    }
    
    sheet.appendRow([
      id,                 // ID
      name,               // Họ và tên
      email,              // Email
      department,         // Phòng ban
      role,               // Chức vụ
      startDate,          // Ngày bắt đầu
      endDate,            // Ngày kết thúc
      leaveType,          // Loại nghỉ
      leaveSession,       // Buổi nghỉ
      reason,             // Lý do nghỉ
      fileUrl,            // Tệp
      status,             // Trạng thái
      note                // Ghi chú
    ]);
    
    return "Đơn nghỉ phép đã được gửi thành công.";
  } catch (error) {
    console.error("Lỗi khi thêm đơn nghỉ phép: " + error.toString());
    throw new Error("Đã xảy ra lỗi: " + error.toString());
  }
}

function checkLeaveOverlap(email, startDate, endDate, leaveType, leaveSession) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pendingSheet = ss.getSheetByName('Đang xử lý');
    var approvedSheet = ss.getSheetByName('Phê duyệt');
    
    // Chuyển thành đối tượng Date cho dễ so sánh
    var newStartDate = new Date(startDate);
    var newEndDate = new Date(endDate);
    
    // Thiết lập ngày bắt đầu và kết thúc về 00:00:00 để so sánh chính xác theo ngày
    newStartDate.setHours(0, 0, 0, 0);
    newEndDate.setHours(0, 0, 0, 0);
    
    // Lấy dữ liệu từ các sheet
    var pendingData = pendingSheet.getDataRange().getValues();
    var approvedData = approvedSheet.getDataRange().getValues();
    
    // Hàm kiểm tra trùng lặp giữa các khoảng thời gian
    function hasDateOverlap(existingStartDate, existingEndDate, newStartDate, newEndDate) {
      // Chuyển đổi thành đối tượng Date nếu cần
      if (!(existingStartDate instanceof Date)) {
        existingStartDate = new Date(existingStartDate);
      }
      if (!(existingEndDate instanceof Date)) {
        existingEndDate = new Date(existingEndDate);
      }
      
      // Đảm bảo tất cả các ngày đều có giờ là 00:00:00
      existingStartDate.setHours(0, 0, 0, 0);
      existingEndDate.setHours(0, 0, 0, 0);
      
      // Kiểm tra trùng lặp
      // Trùng lặp xảy ra khi: 
      // - Thời gian mới bắt đầu trước khi thời gian cũ kết thúc VÀ
      // - Thời gian mới kết thúc sau khi thời gian cũ bắt đầu
      return (newStartDate <= existingEndDate) && (newEndDate >= existingStartDate);
    }
    
    // Hàm kiểm tra trùng lặp trong trường hợp buổi nghỉ
    function hasSessionOverlap(existingSession, newSession, dateOverlap) {
      // Nếu không có trùng lặp về ngày, không cần kiểm tra buổi
      if (!dateOverlap) return false;
      
      // Nếu có một buổi là "Cả ngày", luôn có trùng lặp
      if (existingSession === "Cả ngày" || newSession === "Cả ngày") {
        return true;
      }
      
      // Nếu cả hai buổi giống nhau (Buổi sáng-Buổi sáng, Buổi chiều-Buổi chiều)
      return existingSession === newSession;
    }
    
    // Kiểm tra trùng lặp với các đơn đang xử lý
    for (var i = 1; i < pendingData.length; i++) {
      // Chỉ kiểm tra các đơn của cùng một người
      if (pendingData[i][2] === email) {
        var existingStartDate = parseDate(pendingData[i][5]); // Ngày bắt đầu (column F)
        var existingEndDate = parseDate(pendingData[i][6]);   // Ngày kết thúc (column G)
        var existingLeaveType = pendingData[i][7];            // Loại nghỉ (column H)
        var existingLeaveSession = pendingData[i][8];         // Buổi nghỉ (column I)
        
        // Kiểm tra trùng lặp về ngày
        var dateOverlap = hasDateOverlap(existingStartDate, existingEndDate, newStartDate, newEndDate);
        
        // Nếu có trùng lặp về ngày, kiểm tra thêm về buổi nghỉ
        if (dateOverlap) {
          // Nếu một trong hai là "Từ ngày đến ngày", luôn có trùng lặp
          if (existingLeaveType === "Từ ngày đến ngày" || leaveType === "Từ ngày đến ngày") {
            return true;
          }
          
          // Nếu cả hai đều là "Trong ngày", kiểm tra buổi nghỉ
          if (existingLeaveType === "Trong ngày" && leaveType === "Trong ngày") {
            if (hasSessionOverlap(existingLeaveSession, leaveSession, dateOverlap)) {
              return true;
            }
          }
        }
      }
    }
    
    // Kiểm tra trùng lặp với các đơn đã được phê duyệt (tương tự như trên)
    for (var i = 1; i < approvedData.length; i++) {
      if (approvedData[i][2] === email) {
        var existingStartDate = parseDate(approvedData[i][5]);
        var existingEndDate = parseDate(approvedData[i][6]);
        var existingLeaveType = approvedData[i][7];
        var existingLeaveSession = approvedData[i][8];
        
        var dateOverlap = hasDateOverlap(existingStartDate, existingEndDate, newStartDate, newEndDate);
        
        if (dateOverlap) {
          if (existingLeaveType === "Từ ngày đến ngày" || leaveType === "Từ ngày đến ngày") {
            return true;
          }
          
          if (existingLeaveType === "Trong ngày" && leaveType === "Trong ngày") {
            if (hasSessionOverlap(existingLeaveSession, leaveSession, dateOverlap)) {
              return true;
            }
          }
        }
      }
    }
    
    // Nếu không có trùng lặp nào
    return false;
  } catch (error) {
    console.error("Lỗi khi kiểm tra trùng lặp: " + error.toString());
    // Mặc định trả về false nếu có lỗi, để không chặn việc tạo đơn
    return false;
  }
}

// Sửa hàm editLeaveRequest để kiểm tra đúng số phép còn lại sau ngày 31/3
function editLeaveRequest(form) {
  try {
    var id = form.editId;
    var name = form.editName;
    var email = form.editEmail;
    var department = form.editDepartment;
    var role = form.editRole;
    var startDate = Utilities.formatDate(new Date(form.editStartDate), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var endDate = form.editEndDate ? Utilities.formatDate(new Date(form.editEndDate), Session.getScriptTimeZone(), "dd/MM/yyyy") : startDate;
    var leaveType = form.editLeaveType;
    var leaveSession = form.editLeaveSession || "Cả ngày";
    var reason = form.editReason;
    var fileUrl = form.editFile;
    var status = form.editStatus;
    var note = form.editNote;
    var currentUserEmail = form.currentUserEmail;
    
    // Nếu đang thay đổi trạng thái sang Phê duyệt hoặc Huỷ bỏ
    if (status === 'Phê duyệt' || status === 'Huỷ bỏ') {
      // Kiểm tra quyền phê duyệt (giữ nguyên)
      if (currentUserEmail === email && form.currentUserRole !== 'admin') {
        throw new Error("Bạn không có quyền tự phê duyệt hoặc huỷ bỏ đơn của chính mình!");
      }
      
      var isAdmin = form.currentUserRole === 'admin';
      var isApprover = false;
      
      var userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User');
      var userData = userSheet.getDataRange().getValues();
      
      for (var i = 1; i < userData.length; i++) {
        if (userData[i][2] === email && userData[i][7] === currentUserEmail) {
          isApprover = true;
          break;
        }
      }
      
      if (!isAdmin && !isApprover) {
        throw new Error("Bạn không phải là người phê duyệt được chỉ định cho nhân viên này!");
      }
    }
    
    // Lấy dữ liệu hiện tại để so sánh thay đổi
    var currentLeaveData = getLeaveRequestById(id);
    
    // Nếu không tìm thấy đơn nghỉ phép
    if (!currentLeaveData) {
      throw new Error("Không tìm thấy đơn nghỉ phép với ID này.");
    }
    
    // Kiểm tra nếu có sự thay đổi về ngày nghỉ, buổi nghỉ hoặc loại nghỉ
    var hasDateChanged = 
      startDate !== currentLeaveData[5] || 
      endDate !== currentLeaveData[6] || 
      leaveType !== currentLeaveData[7] || 
      leaveSession !== currentLeaveData[8];
    
    // Nếu có thay đổi về ngày/buổi nghỉ và không phải admin đang thay đổi trạng thái, kiểm tra trùng lặp
    if (hasDateChanged && !(form.currentUserRole === 'admin' && (status === 'Phê duyệt' || status === 'Huỷ bỏ'))) {
      // Kiểm tra trùng lặp, loại trừ đơn hiện tại đang sửa
      var overlapWithOthers = checkLeaveOverlapExcludingSelf(
        email, 
        form.editStartDate, 
        form.editEndDate || form.editStartDate, 
        leaveType, 
        leaveSession, 
        id
      );
      
      if (overlapWithOthers) {
        throw new Error("Bạn đã có đơn nghỉ phép cho thời gian này. Vui lòng kiểm tra lại!");
      }
    }
    
    // Nếu đơn đang ở trạng thái xử lý và có thay đổi về ngày, kiểm tra số ngày nghỉ
    if (status === 'Đang xử lý' && hasDateChanged) {
      // Tính số ngày nghỉ mới
      var leaveDays = calculateLeaveDays(
        form.editStartDate,
        form.editEndDate || form.editStartDate,
        leaveType,
        leaveSession
      );
      
      // Lấy số ngày nghỉ hiện tại của đơn này để không tính trùng
      var currentLeaveDays = calculateLeaveDays(
        currentLeaveData[5], // startDate
        currentLeaveData[6], // endDate
        currentLeaveData[7], // leaveType
        currentLeaveData[8]  // leaveSession
      );
      
      // Lấy ngày bắt đầu và kết thúc nghỉ để kiểm tra
      var startDateObj = new Date(form.editStartDate);
      var endDateObj = new Date(form.editEndDate || form.editStartDate);
      
      // Lấy ngày hết hạn phép năm trước (31/03 năm hiện tại)
      var today = new Date();
      var expireYear = today.getFullYear();
      var expireDate = new Date(expireYear, 2, 31, 23, 59, 59); // 31/3 năm hiện tại, 23:59:59
      
      // Lấy số phép đã sử dụng và còn lại
      var leaveBalance = calculateLeaveBalance(email);
      
      // Xử lý đặc biệt cho trường hợp khoảng thời gian chuyển tiếp (từ trước 31/3 đến sau 31/3)
      if (leaveType === 'Từ ngày đến ngày' && startDateObj <= expireDate && endDateObj > expireDate) {
        // Đây là trường hợp khoảng thời gian chuyển tiếp
        // Tính số ngày trước hoặc vào 31/3
        var tempEndDate = new Date(expireYear, 2, 31); // 31/3 của năm hiện tại
        var daysBeforeExpire = calculateLeaveDays(
          startDateObj, 
          tempEndDate, 
          "Từ ngày đến ngày", 
          "Cả ngày"
        );
        
        // Tính số ngày sau 31/3
        var tempStartDate = new Date(expireYear, 3, 1); // 1/4 của năm hiện tại
        var daysAfterExpire = calculateLeaveDays(
          tempStartDate, 
          endDateObj, 
          "Từ ngày đến ngày", 
          "Cả ngày"
        );
        
        console.log("Sửa đơn - Trường hợp chuyển tiếp: " + daysBeforeExpire + " ngày trước 31/3 và " + 
                  daysAfterExpire + " ngày sau 31/3");
        
        // Kiểm tra phép năm trước (đủ cho phần trước 31/3 không)
        // Cần tính toán lại số phép đã sử dụng và còn lại dựa trên đơn hiện tại
        var adjustedPreviousYearRemaining = leaveBalance.remainingPreviousYear;
        var adjustedCurrentYearRemaining = leaveBalance.remainingCurrentYear;
        
        // Nếu đơn hiện tại có một phần trước 31/3, cộng lại vào phép năm trước còn lại
        if (new Date(currentLeaveData[5]) <= expireDate) {
          var currentDaysBeforeExpire = calculateLeaveDays(
            new Date(currentLeaveData[5]),
            new Date(expireYear, 2, 31) < new Date(currentLeaveData[6]) ? new Date(expireYear, 2, 31) : new Date(currentLeaveData[6]),
            currentLeaveData[7],
            currentLeaveData[8]
          );
          
          if (currentDaysBeforeExpire > 0 && currentDaysBeforeExpire <= leaveBalance.usedPreviousYear) {
            adjustedPreviousYearRemaining += currentDaysBeforeExpire;
          }
        }
        
        // Nếu đơn hiện tại có một phần sau 31/3, cộng lại vào phép năm nay còn lại
        if (new Date(currentLeaveData[6]) > expireDate) {
          var currentDaysAfterExpire = calculateLeaveDays(
            new Date(currentLeaveData[5]) > new Date(expireYear, 3, 1) ? new Date(currentLeaveData[5]) : new Date(expireYear, 3, 1),
            new Date(currentLeaveData[6]),
            currentLeaveData[7],
            currentLeaveData[8]
          );
          
          if (currentDaysAfterExpire > 0) {
            adjustedCurrentYearRemaining += currentDaysAfterExpire;
          }
        }
        
        // Kiểm tra phép năm trước cho phần trước 31/3
        if (daysBeforeExpire > adjustedPreviousYearRemaining) {
          throw new Error("Số ngày nghỉ trước 31/3 (" + daysBeforeExpire + 
                         " ngày) vượt quá số phép năm trước còn lại (" + 
                         adjustedPreviousYearRemaining + " ngày).");
        }
        
        // Kiểm tra phép năm nay cho phần sau 31/3
        if (daysAfterExpire > adjustedCurrentYearRemaining) {
          throw new Error("Số ngày nghỉ sau 31/3 (" + daysAfterExpire + 
                         " ngày) vượt quá số phép năm nay còn lại (" + 
                         adjustedCurrentYearRemaining + " ngày).");
        }
      } else {
        // Trường hợp bình thường (không nằm ở khoảng thời gian chuyển tiếp)
        // Điều chỉnh số phép còn lại dựa trên đơn hiện tại
        var adjustedRemaining;
        
        // Đảm bảo so sánh chính xác với ngày 31/3 (làm tròn xuống 00:00:00)
        var compareStartDate = new Date(startDateObj.getFullYear(), startDateObj.getMonth(), startDateObj.getDate());
        var compareExpireDate = new Date(expireDate.getFullYear(), expireDate.getMonth(), expireDate.getDate());
        
        if (compareStartDate > compareExpireDate) {
          // Nếu nghỉ sau 31/3, chỉ tính phép năm nay
          // Cần điều chỉnh số phép năm nay nếu đơn hiện tại cũng sử dụng phép năm nay
          adjustedRemaining = leaveBalance.remainingCurrentYear;
          
          // Nếu đơn hiện tại cũng sử dụng phép năm nay, cộng lại số ngày đó
          if (new Date(currentLeaveData[5]) > expireDate || 
             (new Date(currentLeaveData[5]) <= expireDate && leaveBalance.usedPreviousYear < currentLeaveDays)) {
            var currentDaysFromCurrentYear = Math.min(currentLeaveDays, leaveBalance.used);
            adjustedRemaining += currentDaysFromCurrentYear;
          }
          
          if (leaveDays > adjustedRemaining) {
            throw new Error("Số ngày nghỉ (" + leaveDays + 
                          " ngày) vượt quá số phép năm nay còn lại (" + 
                          adjustedRemaining + " ngày).");
          }
        } else {
          // Nếu nghỉ trước hoặc vào 31/3, ưu tiên dùng phép năm trước
          // Điều chỉnh cả phép năm trước và tổng phép
          var adjustedPreviousYearRemaining = leaveBalance.remainingPreviousYear;
          
          // Nếu đơn hiện tại sử dụng phép năm trước, cộng lại số ngày đó
          if (new Date(currentLeaveData[5]) <= expireDate) {
            var currentDaysFromPreviousYear = Math.min(currentLeaveDays, leaveBalance.usedPreviousYear);
            adjustedPreviousYearRemaining += currentDaysFromPreviousYear;
          }
          
          // Tính toán tổng phép điều chỉnh
          var adjustedTotalRemaining = adjustedPreviousYearRemaining + leaveBalance.remainingCurrentYear;
          
          if (leaveDays > adjustedPreviousYearRemaining) {
            // Nếu phép năm trước không đủ, kiểm tra tổng phép
            if (leaveDays > adjustedTotalRemaining) {
              throw new Error("Số ngày nghỉ (" + leaveDays + 
                            " ngày) vượt quá tổng số phép còn lại (" + 
                            adjustedTotalRemaining + " ngày).");
            }
          }
        }
      }
    }
    
    // Thực hiện cập nhật đơn nghỉ phép
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pendingSheet = ss.getSheetByName('Đang xử lý');
    var approvedSheet = ss.getSheetByName('Phê duyệt');
    var disapprovedSheet = ss.getSheetByName('Huỷ bỏ');
    var dataRange = pendingSheet.getDataRange();
    var values = dataRange.getValues();

    for (var i = 1; i < values.length; i++) {
      if (values[i][0] == id) {
        var rowData = values[i].slice();
        if (rowData[11] !== status) { // Status index is 11
          var targetSheet;
          if (status === 'Phê duyệt') {
            targetSheet = approvedSheet;
          } else if (status === 'Huỷ bỏ') {
            targetSheet = disapprovedSheet;
          } else {
            throw new Error("Trạng thái không hợp lệ. Phải là 'Phê duyệt' hoặc 'Huỷ bỏ'.");
          }
          
          targetSheet.appendRow([
            id, name, email, department, role, startDate, endDate, 
            leaveType, leaveSession, reason, fileUrl, status, note
          ]);
          
          pendingSheet.deleteRow(i + 1);
          return "Đơn nghỉ phép đã được cập nhật thành công.";
        } else {
          rowData[1] = name;
          rowData[2] = email;
          rowData[3] = department;
          rowData[4] = role;
          rowData[5] = startDate;
          rowData[6] = endDate;
          rowData[7] = leaveType;
          rowData[8] = leaveSession;
          rowData[9] = reason;
          rowData[10] = fileUrl;
          rowData[12] = note;
          
          pendingSheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
          return "Đơn nghỉ phép đã được cập nhật thành công.";
        }
      }
    }
    throw new Error("Không tìm thấy đơn nghỉ phép với ID này.");
  } catch (error) {
    throw new Error("Lỗi: " + error.toString());
  }
}

function checkLeaveOverlapExcludingSelf(email, startDate, endDate, leaveType, leaveSession, excludeId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pendingSheet = ss.getSheetByName('Đang xử lý');
    var approvedSheet = ss.getSheetByName('Phê duyệt');
    
    // Chuyển thành đối tượng Date cho dễ so sánh
    var newStartDate = new Date(startDate);
    var newEndDate = new Date(endDate);
    
    // Thiết lập ngày bắt đầu và kết thúc về 00:00:00 để so sánh chính xác theo ngày
    newStartDate.setHours(0, 0, 0, 0);
    newEndDate.setHours(0, 0, 0, 0);
    
    // Lấy dữ liệu từ các sheet
    var pendingData = pendingSheet.getDataRange().getValues();
    var approvedData = approvedSheet.getDataRange().getValues();
    
    // Hàm kiểm tra trùng lặp giữa các khoảng thời gian
    function hasDateOverlap(existingStartDate, existingEndDate, newStartDate, newEndDate) {
      // Chuyển đổi thành đối tượng Date nếu cần
      if (!(existingStartDate instanceof Date)) {
        existingStartDate = new Date(existingStartDate);
      }
      if (!(existingEndDate instanceof Date)) {
        existingEndDate = new Date(existingEndDate);
      }
      
      // Đảm bảo tất cả các ngày đều có giờ là 00:00:00
      existingStartDate.setHours(0, 0, 0, 0);
      existingEndDate.setHours(0, 0, 0, 0);
      
      // Kiểm tra trùng lặp
      return (newStartDate <= existingEndDate) && (newEndDate >= existingStartDate);
    }
    
    // Hàm kiểm tra trùng lặp trong trường hợp buổi nghỉ
    function hasSessionOverlap(existingSession, newSession, dateOverlap) {
      // Nếu không có trùng lặp về ngày, không cần kiểm tra buổi
      if (!dateOverlap) return false;
      
      // Nếu có một buổi là "Cả ngày", luôn có trùng lặp
      if (existingSession === "Cả ngày" || newSession === "Cả ngày") {
        return true;
      }
      
      // Nếu cả hai buổi giống nhau (Buổi sáng-Buổi sáng, Buổi chiều-Buổi chiều)
      return existingSession === newSession;
    }
    
    // Kiểm tra trùng lặp với các đơn đang xử lý
    for (var i = 1; i < pendingData.length; i++) {
      // Chỉ kiểm tra các đơn của cùng một người và không phải đơn đang sửa
      if (pendingData[i][2] === email && pendingData[i][0] != excludeId) {
        var existingStartDate = parseDate(pendingData[i][5]); // Ngày bắt đầu (column F)
        var existingEndDate = parseDate(pendingData[i][6]);   // Ngày kết thúc (column G)
        var existingLeaveType = pendingData[i][7];            // Loại nghỉ (column H)
        var existingLeaveSession = pendingData[i][8];         // Buổi nghỉ (column I)
        
        // Kiểm tra trùng lặp về ngày
        var dateOverlap = hasDateOverlap(existingStartDate, existingEndDate, newStartDate, newEndDate);
        
        // Nếu có trùng lặp về ngày, kiểm tra thêm về buổi nghỉ
        if (dateOverlap) {
          // Nếu một trong hai là "Từ ngày đến ngày", luôn có trùng lặp
          if (existingLeaveType === "Từ ngày đến ngày" || leaveType === "Từ ngày đến ngày") {
            return true;
          }
          
          // Nếu cả hai đều là "Trong ngày", kiểm tra buổi nghỉ
          if (existingLeaveType === "Trong ngày" && leaveType === "Trong ngày") {
            if (hasSessionOverlap(existingLeaveSession, leaveSession, dateOverlap)) {
              return true;
            }
          }
        }
      }
    }
    
    // Kiểm tra trùng lặp với các đơn đã được phê duyệt (tương tự như trên)
    for (var i = 1; i < approvedData.length; i++) {
      if (approvedData[i][2] === email) {
        var existingStartDate = parseDate(approvedData[i][5]);
        var existingEndDate = parseDate(approvedData[i][6]);
        var existingLeaveType = approvedData[i][7];
        var existingLeaveSession = approvedData[i][8];
        
        var dateOverlap = hasDateOverlap(existingStartDate, existingEndDate, newStartDate, newEndDate);
        
        if (dateOverlap) {
          if (existingLeaveType === "Từ ngày đến ngày" || leaveType === "Từ ngày đến ngày") {
            return true;
          }
          
          if (existingLeaveType === "Trong ngày" && leaveType === "Trong ngày") {
            if (hasSessionOverlap(existingLeaveSession, leaveSession, dateOverlap)) {
              return true;
            }
          }
        }
      }
    }
    
    // Nếu không có trùng lặp nào
    return false;
  } catch (error) {
    console.error("Lỗi khi kiểm tra trùng lặp: " + error.toString());
    // Mặc định trả về false nếu có lỗi, để không chặn việc sửa đơn
    return false;
  }
}

function getLeaveRequestById(id) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Đang xử lý");
    if (!sheet) {
      console.error("Không tìm thấy sheet Đang xử lý");
      return null;
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      console.error("Sheet không có dữ liệu");
      return null;
    }
    
    // Lấy tất cả dữ liệu ID từ cột A để debug
    var allIds = sheet.getRange("A2:A" + lastRow).getValues();
    console.log("Tất cả IDs trong sheet:", JSON.stringify(allIds));
    console.log("ID cần tìm:", id, "Kiểu:", typeof id);
    
    // Lấy tất cả dữ liệu
    var data = sheet.getRange("A2:M" + lastRow).getValues();
    
    for (var i = 0; i < data.length; i++) {
      // Lấy ID từ sheet và chuyển thành chuỗi
      var rowId = data[i][0];
      var rowIdStr = String(rowId).replace(/^['"]|['"]$/g, ""); // Loại bỏ dấu nháy đơn hoặc kép ở đầu và cuối
      var searchIdStr = String(id).replace(/^['"]|['"]$/g, "");
      
      console.log("So sánh: [" + rowIdStr + "] với [" + searchIdStr + "]");
      
      // So sánh cả hai dạng: nguyên bản và sau khi chuyển đổi
      if (rowId == id || rowIdStr === searchIdStr) {
        console.log("Đã tìm thấy dữ liệu");
        return data[i];
      }
    }
    
    console.error("Không tìm thấy dữ liệu với ID:", id);
    return null;
  } catch (error) {
    console.error("Lỗi khi lấy dữ liệu:", error.message, error.stack);
    return null;
  }
}

function deleteLeaveRequest(id) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Đang xử lý');
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var found = false;
    
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] == id) { 
        var fileUrl = values[i][10]; // Vị trí cột Tệp đã thay đổi sang cột 11 (index 10)
        if (fileUrl) {
          var fileId = getIdFromUrl(fileUrl);
          if (fileId) {
            try {
              DriveApp.getFileById(fileId).setTrashed(true);
            } catch (e) {
              // Bỏ qua lỗi nếu không thể xóa file (có thể file đã bị xóa)
              console.error("Không thể xóa file: " + e.toString());
            }
          }
        }
        sheet.deleteRow(i + 1);
        found = true;
        break;
      }
    }
    
    if (!found) {
      throw new Error("Không tìm thấy đơn nghỉ phép với ID này.");
    }
    
    return "Đơn nghỉ phép và tệp đã xóa vĩnh viễn.";
    
  } catch (error) {
    throw new Error("Lỗi: " + error.toString());
  }
}

function getIdFromUrl(url) {
  if (!url) return null;
  var match = /\/d\/([^\/]+)/.exec(url);
  return match && match[1];
}

function getDepartments() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) {
    // Tạo sheet Config nếu không tồn tại
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Config');
    sheet.getRange('A1').setValue('Phòng ban');
    sheet.getRange('B1').setValue('Chức vụ');
    sheet.getRange('C1').setValue('Ngày nghỉ lễ');
  }
  
  var lastRow = Math.max(sheet.getLastRow(), 1);
  // Bỏ qua dòng tiêu đề
  if (lastRow <= 1) return [];
  
  var data = sheet.getRange('A2:A' + lastRow).getValues();
  var departments = [];
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] !== "") {
      departments.push(data[i][0]);
    }
  }
  
  return departments;
}

function getRoles() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) return [];
  
  var lastRow = Math.max(sheet.getLastRow(), 1);
  // Bỏ qua dòng tiêu đề
  if (lastRow <= 1) return [];
  
  var data = sheet.getRange('B2:B' + lastRow).getValues();
  var roles = [];
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] !== "") {
      roles.push(data[i][0]);
    }
  }
  
  return roles;
}

// Hàm lấy dữ liệu Config
function getConfigData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) {
    // Tạo sheet Config nếu không tồn tại
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Config');
    sheet.getRange('A1').setValue('Phòng ban');
    sheet.getRange('B1').setValue('Chức vụ');
    sheet.getRange('C1').setValue('Ngày nghỉ lễ');
    return {departments: [], positions: [], holidays: []};
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return {departments: [], positions: [], holidays: []};
  }
  
  var data = sheet.getRange(2, 1, lastRow-1, 3).getValues();
  var departments = [];
  var positions = [];
  var holidays = [];
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] !== "") departments.push(data[i][0]);
    if (data[i][1] !== "") positions.push(data[i][1]);
    if (data[i][2] !== "") {
      // Nếu dữ liệu là Date, định dạng thành chuỗi ngày tháng
      if (data[i][2] instanceof Date) {
        var dateStr = Utilities.formatDate(data[i][2], Session.getScriptTimeZone(), "dd/MM/yyyy");
        holidays.push(dateStr);
      } else {
        holidays.push(data[i][2]);
      }
    }
  }
  
  return {departments: departments, positions: positions, holidays: holidays};
}

// Hàm thêm phòng ban mới
function addDepartment(name) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Config');
      sheet.getRange('A1').setValue('Phòng ban');
      sheet.getRange('B1').setValue('Chức vụ');
      sheet.getRange('C1').setValue('Ngày nghỉ lễ');
    }
    
    if (checkDepartmentExists(name)) {
      throw new Error("Phòng ban này đã tồn tại!");
    }
    
    var lastRow = Math.max(sheet.getLastRow(), 1);
    sheet.getRange('A' + (lastRow + 1)).setValue(name);
    return "Phòng ban đã được thêm thành công.";
  } catch (error) {
    throw new Error("Lỗi khi thêm phòng ban: " + error.toString());
  }
}

// Hàm thêm chức vụ mới
function addRole(name) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Config');
      sheet.getRange('A1').setValue('Phòng ban');
      sheet.getRange('B1').setValue('Chức vụ');
      sheet.getRange('C1').setValue('Ngày nghỉ lễ');
    }
    
    if (checkRoleExists(name)) {
      throw new Error("Chức vụ này đã tồn tại!");
    }
    
    var lastRow = Math.max(sheet.getLastRow(), 1);
    sheet.getRange('B' + (lastRow + 1)).setValue(name);
    return "Chức vụ đã được thêm thành công.";
  } catch (error) {
    throw new Error("Lỗi khi thêm chức vụ: " + error.toString());
  }
}

// Hàm thêm ngày nghỉ lễ mới
function addHoliday(date) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Config');
      sheet.getRange('A1').setValue('Phòng ban');
      sheet.getRange('B1').setValue('Chức vụ');
      sheet.getRange('C1').setValue('Ngày nghỉ lễ');
    }
    
    // Kiểm tra định dạng ngày
    var holidayDate;
    try {
      holidayDate = new Date(date);
      if (isNaN(holidayDate.getTime())) {
        throw new Error("Ngày không hợp lệ");
      }
    } catch (e) {
      throw new Error("Định dạng ngày không hợp lệ");
    }
    
    // Kiểm tra trùng lặp
    if (checkHolidayExists(date)) {
      throw new Error("Ngày nghỉ lễ này đã tồn tại!");
    }
    
    var lastRow = Math.max(sheet.getLastRow(), 1);
    sheet.getRange('C' + (lastRow + 1)).setValue(holidayDate);
    return "Ngày nghỉ lễ đã được thêm thành công.";
  } catch (error) {
    throw new Error("Lỗi khi thêm ngày nghỉ lễ: " + error.toString());
  }
}

// Hàm xóa phòng ban
function deleteDepartment(name) {
  try {
    if (checkDepartmentInUse(name)) {
      throw new Error("Không thể xóa phòng ban đã được sử dụng trong dữ liệu!");
    }
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!sheet) {
      throw new Error("Không tìm thấy sheet Config.");
    }
    
    var found = false;
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues();
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === name) {
        sheet.getRange(i + 2, 1).setValue(""); // Xóa giá trị
        found = true;
        break;
      }
    }
    
    if (!found) {
      throw new Error("Không tìm thấy phòng ban này.");
    }
    
    return "Phòng ban đã được xóa thành công.";
  } catch (error) {
    throw new Error("Lỗi khi xóa phòng ban: " + error.toString());
  }
}

// Hàm xóa chức vụ
function deleteRole(name) {
  try {
    if (checkRoleInUse(name)) {
      throw new Error("Không thể xóa chức vụ đã được sử dụng trong dữ liệu!");
    }
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!sheet) {
      throw new Error("Không tìm thấy sheet Config.");
    }
    
    var found = false;
    var data = sheet.getRange(2, 2, sheet.getLastRow()-1, 1).getValues();
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === name) {
        sheet.getRange(i + 2, 2).setValue(""); // Xóa giá trị
        found = true;
        break;
      }
    }
    
    if (!found) {
      throw new Error("Không tìm thấy chức vụ này.");
    }
    
    return "Chức vụ đã được xóa thành công.";
  } catch (error) {
    throw new Error("Lỗi khi xóa chức vụ: " + error.toString());
  }
}

// Hàm xóa ngày nghỉ lễ
function deleteHoliday(date) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!sheet) {
      throw new Error("Không tìm thấy sheet Config.");
    }
    
    var found = false;
    var data = sheet.getRange(2, 3, sheet.getLastRow()-1, 1).getValues();
    
    // Cải tiến: Không cần chuyển đổi giá trị ngày đầu vào thành đối tượng Date
    // Thay vào đó, so sánh trực tiếp chuỗi
    var targetDateStr = date;
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] instanceof Date) {
        // Nếu giá trị trong ô là Date, chuyển đổi thành chuỗi định dạng dd/MM/yyyy
        var dateStr = Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), "dd/MM/yyyy");
        if (dateStr === targetDateStr) {
          sheet.getRange(i + 2, 3).setValue(""); // Xóa giá trị
          found = true;
          break;
        }
      } else if (data[i][0] === targetDateStr) {
        // Nếu giá trị trong ô là chuỗi, so sánh trực tiếp
        sheet.getRange(i + 2, 3).setValue(""); // Xóa giá trị
        found = true;
        break;
      }
    }
    
    if (!found) {
      throw new Error("Không tìm thấy ngày nghỉ lễ này.");
    }
    
    return "Ngày nghỉ lễ đã được xóa thành công.";
  } catch (error) {
    throw new Error("Lỗi khi xóa ngày nghỉ lễ: " + error.toString());
  }
}

// Hàm sửa phòng ban
function editDepartment(oldName, newName) {
  try {
    if (checkDepartmentInUse(oldName)) {
      throw new Error("Không thể sửa phòng ban đã được sử dụng trong dữ liệu!");
    }
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!sheet) {
      throw new Error("Không tìm thấy sheet Config.");
    }
    
    var found = false;
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues();
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === oldName) {
        sheet.getRange(i + 2, 1).setValue(newName);
        found = true;
        break;
      }
    }
    
    if (!found) {
      throw new Error("Không tìm thấy phòng ban này.");
    }
    
    return "Phòng ban đã được cập nhật thành công.";
  } catch (error) {
    throw new Error("Lỗi khi cập nhật phòng ban: " + error.toString());
  }
}

// Hàm sửa chức vụ
function editRole(oldName, newName) {
  try {
    if (checkRoleInUse(oldName)) {
      throw new Error("Không thể sửa chức vụ đã được sử dụng trong dữ liệu!");
    }
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!sheet) {
      throw new Error("Không tìm thấy sheet Config.");
    }
    
    var found = false;
    var data = sheet.getRange(2, 2, sheet.getLastRow()-1, 1).getValues();
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === oldName) {
        sheet.getRange(i + 2, 2).setValue(newName);
        found = true;
        break;
      }
    }
    
    if (!found) {
      throw new Error("Không tìm thấy chức vụ này.");
    }
    
    return "Chức vụ đã được cập nhật thành công.";
  } catch (error) {
    throw new Error("Lỗi khi cập nhật chức vụ: " + error.toString());
  }
}

function getUserData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User');
    if (!sheet) {
      throw new Error("Không tìm thấy sheet 'User'");
    }
    const dataRange = sheet.getRange('A2:K' + sheet.getLastRow()); // Mở rộng phạm vi đến cột K
    const values = dataRange.getValues();
    const filteredValues = values.filter(row => row[0] !== '');
    return JSON.stringify(filteredValues);
  } catch (error) {
    console.error("Lỗi khi lấy dữ liệu từ User:", error.message);
    return JSON.stringify([]);
  }
}

function getPendingData(email, isAdmin) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Đang xử lý');
    if (!sheet || sheet.getLastRow() <= 1) {
      return JSON.stringify([]);
    }
    
    const dataRange = sheet.getRange('A2:M' + sheet.getLastRow());
    const values = dataRange.getValues();
    let filteredValues;
    
    // Lấy thông tin người dùng
    var user = getUserByUsername(email);
    var isApprover = false;
    var approverFor = [];
    
    // Kiểm tra nếu người dùng là người phê duyệt
    if (user && user.role !== 'Admin') {
      var userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User');
      var userData = userSheet.getDataRange().getValues();
      
      for (var i = 1; i < userData.length; i++) {
        if (userData[i][7] === email) { // Email người phê duyệt ở cột 8 (index 7)
          isApprover = true;
          approverFor.push(userData[i][2]); // Email của người được phê duyệt
        }
      }
    }
    
    if (isAdmin === 'admin') {
      // Admin thấy tất cả dữ liệu
      filteredValues = values.filter(row => row[0] !== '');
    } else if (isApprover) {
      // Người phê duyệt thấy cả dữ liệu của họ và của người họ phê duyệt
      filteredValues = values.filter(row => 
        row[0] !== '' && (row[2] === email || approverFor.includes(row[2]))
      );
    } else {
      // User chỉ thấy dữ liệu của mình
      filteredValues = values.filter(row => row[0] !== '' && row[2] === email);
    }
    
    return JSON.stringify(filteredValues);
  } catch (error) {
    console.error("Lỗi khi lấy dữ liệu Đang xử lý:", error.message);
    return JSON.stringify([]);
  }
}

function getApprovedData(email, isAdmin) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Phê duyệt');
    if (!sheet || sheet.getLastRow() <= 1) {
      return JSON.stringify([]);
    }
    
    const dataRange = sheet.getRange('A2:M' + sheet.getLastRow());
    const values = dataRange.getValues();
    let filteredValues;
    
    // Lấy thông tin người dùng
    var user = getUserByUsername(email);
    var isApprover = false;
    var approverFor = [];
    
    // Kiểm tra nếu người dùng là người phê duyệt
    if (user && user.role !== 'Admin') {
      var userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User');
      var userData = userSheet.getDataRange().getValues();
      
      for (var i = 1; i < userData.length; i++) {
        if (userData[i][7] === email) { // Email người phê duyệt ở cột 8 (index 7)
          isApprover = true;
          approverFor.push(userData[i][2]); // Email của người được phê duyệt
        }
      }
    }
    
    if (isAdmin === 'admin') {
      // Admin thấy tất cả dữ liệu
      filteredValues = values.filter(row => row[0] !== '');
    } else if (isApprover) {
      // Người phê duyệt thấy cả dữ liệu của họ và của người họ phê duyệt
      filteredValues = values.filter(row => 
        row[0] !== '' && (row[2] === email || approverFor.includes(row[2]))
      );
    } else {
      // User chỉ thấy dữ liệu của mình
      filteredValues = values.filter(row => row[0] !== '' && row[2] === email);
    }
    
    return JSON.stringify(filteredValues);
  } catch (error) {
    console.error("Lỗi khi lấy dữ liệu Phê duyệt:", error.message);
    return JSON.stringify([]);
  }
}

function getDisapprovedData(email, isAdmin) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Huỷ bỏ');
    if (!sheet || sheet.getLastRow() <= 1) {
      return JSON.stringify([]);
    }
    
    const dataRange = sheet.getRange('A2:M' + sheet.getLastRow());
    const values = dataRange.getValues();
    let filteredValues;
    
    // Lấy thông tin người dùng
    var user = getUserByUsername(email);
    var isApprover = false;
    var approverFor = [];
    
    // Kiểm tra nếu người dùng là người phê duyệt
    if (user && user.role !== 'Admin') {
      var userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User');
      var userData = userSheet.getDataRange().getValues();
      
      for (var i = 1; i < userData.length; i++) {
        if (userData[i][7] === email) { // Email người phê duyệt ở cột 8 (index 7)
          isApprover = true;
          approverFor.push(userData[i][2]); // Email của người được phê duyệt
        }
      }
    }
    
    if (isAdmin === 'admin') {
      // Admin thấy tất cả dữ liệu
      filteredValues = values.filter(row => row[0] !== '');
    } else if (isApprover) {
      // Người phê duyệt thấy cả dữ liệu của họ và của người họ phê duyệt
      filteredValues = values.filter(row => 
        row[0] !== '' && (row[2] === email || approverFor.includes(row[2]))
      );
    } else {
      // User chỉ thấy dữ liệu của mình
      filteredValues = values.filter(row => row[0] !== '' && row[2] === email);
    }
    
    return JSON.stringify(filteredValues);
  } catch (error) {
    console.error("Lỗi khi lấy dữ liệu Huỷ bỏ:", error.message);
    return JSON.stringify([]);
  }
}

// Kiểm tra tên phòng ban trùng lặp
function checkDepartmentExists(name) {
  var departments = getDepartments();
  name = name.trim().toLowerCase();
  
  for (var i = 0; i < departments.length; i++) {
    if (departments[i].trim().toLowerCase() === name) {
      return true;
    }
  }
  return false;
}

// Kiểm tra tên chức vụ trùng lặp
function checkRoleExists(name) {
  var roles = getRoles();
  name = name.trim().toLowerCase();
  
  for (var i = 0; i < roles.length; i++) {
    if (roles[i].trim().toLowerCase() === name) {
      return true;
    }
  }
  return false;
}

// Kiểm tra ngày nghỉ lễ trùng lặp
function checkHolidayExists(date) {
  var holidays = getHolidays();
  
  // Chuẩn hóa định dạng ngày để so sánh
  var targetDate;
  try {
    targetDate = new Date(date);
    if (isNaN(targetDate.getTime())) {
      return false;
    }
  } catch (e) {
    return false;
  }
  
  var targetDateStr = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  for (var i = 0; i < holidays.length; i++) {
    var holidayDate;
    try {
      if (holidays[i].includes('/')) {
        // Nếu dữ liệu đã ở dạng chuỗi dd/MM/yyyy
        holidayDate = holidays[i];
      } else {
        // Nếu dữ liệu dạng Date, chuyển sang chuỗi
        holidayDate = Utilities.formatDate(new Date(holidays[i]), Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
    } catch (e) {
      continue;
    }
    
    if (holidayDate === targetDateStr) {
      return true;
    }
  }
  return false;
}

// Thêm các hàm để quản lý người dùng
function addUser(form) {
  try {
    var name = form.name;
    var email = form.email;
    var password = form.password;
    var image = form.image;
    var department = form.department;
    var role = form.role;
    var approverEmail = form.approverEmail;
    var leaveStartDate = form.leaveStartDate ? Utilities.formatDate(new Date(form.leaveStartDate), Session.getScriptTimeZone(), "dd/MM/yyyy") : "";
    // Không dùng giá trị totalLeave nữa, để trống để công thức trong sheet tính toán
    var previousYearLeave = parseFloat(form.previousYearLeave) || 0; // Phép năm trước
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('User');
    var lastRow = sheet.getLastRow();
    var id = lastRow; // ID tự tăng
    
    // Kiểm tra xem email đã tồn tại chưa
    var data = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === email) {
        throw new Error("Email đã tồn tại trong hệ thống!");
      }
    }
    
    // Thêm dòng và để trống ô tổng phép năm (cột 10)
    sheet.appendRow([
      id, name, email, password, image, 
      department, role, approverEmail, leaveStartDate, 
      "", previousYearLeave // Cột totalLeave để trống
    ]);
    return "Người dùng đã được thêm thành công.";
  } catch (error) {
    throw new Error("Lỗi khi thêm người dùng: " + error.toString());
  }
}

function editUser(form) {
  try {
    var id = form.id;
    var name = form.name;
    var email = form.email;
    var oldEmail = form.oldEmail; // Email cũ
    var password = form.password;
    var image = form.image;
    var department = form.department;
    var role = form.role;
    var approverEmail = form.approverEmail;
    var leaveStartDate = form.leaveStartDate ? Utilities.formatDate(new Date(form.leaveStartDate), Session.getScriptTimeZone(), "dd/MM/yyyy") : "";
    // Không dùng giá trị totalLeave nữa
    var previousYearLeave = parseFloat(form.previousYearLeave) || 0; // Phép năm trước
    
    Logger.log("Thông tin cập nhật: ID=" + id + ", email mới=" + email + ", email cũ=" + oldEmail);
    
    // Nếu thay đổi email và người dùng có dữ liệu
    if (oldEmail !== email && checkUserHasData(oldEmail)) {
      throw new Error("Không thể thay đổi email của người dùng đã có dữ liệu trong hệ thống!");
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('User');
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] == id) {
        // Cập nhật dữ liệu
        sheet.getRange(i + 1, 2).setValue(name);
        sheet.getRange(i + 1, 3).setValue(email);
        sheet.getRange(i + 1, 4).setValue(password);
        sheet.getRange(i + 1, 5).setValue(image);
        sheet.getRange(i + 1, 6).setValue(department);
        sheet.getRange(i + 1, 7).setValue(role);
        sheet.getRange(i + 1, 8).setValue(approverEmail);
        sheet.getRange(i + 1, 9).setValue(leaveStartDate);
        // Không cập nhật cột 10 (totalLeave) để giữ nguyên công thức
        sheet.getRange(i + 1, 11).setValue(previousYearLeave); // Cập nhật cột phép năm trước
        return "Cập nhật người dùng thành công.";
      }
    }
    throw new Error("Không tìm thấy người dùng với ID này.");
  } catch (error) {
    throw new Error("Lỗi khi cập nhật người dùng: " + error.toString());
  }
}

function deleteUser(id) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('User');
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] == id) {
        var email = values[i][2];
        
        // Kiểm tra người dùng có dữ liệu không
        if (checkUserHasData(email)) {
          throw new Error("Không thể xóa người dùng đã có dữ liệu trong hệ thống!");
        }
        
        sheet.deleteRow(i + 1);
        return "Người dùng đã được xóa thành công.";
      }
    }
    throw new Error("Không tìm thấy người dùng với ID này.");
  } catch (error) {
    throw new Error("Lỗi khi xóa người dùng: " + error.toString());
  }
}

function getUserById(id) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("User");
    if (!sheet) return null;
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return null;
    
    var data = sheet.getRange("A2:K" + lastRow).getValues();
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] == id) {
        // Chuyển đổi đối tượng Date sang chuỗi để tránh lỗi
        var row = data[i].slice(); // Tạo bản sao dữ liệu
        
        // Xử lý cột ngày tính phép (index 8)
        if (row[8] instanceof Date) {
          row[8] = Utilities.formatDate(row[8], Session.getScriptTimeZone(), "dd/MM/yyyy");
        }
        
        // Log dữ liệu để debug
        Logger.log("Dữ liệu người dùng ID " + id + ": " + JSON.stringify(row));
        
        return row;
      }
    }
    
    Logger.log("Không tìm thấy người dùng với ID: " + id);
    return null;
  } catch (error) {
    Logger.log("Lỗi khi lấy dữ liệu người dùng: " + error.message);
    return null;
  }
}

function moveDataToPending(id, sourceSheet) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sourceSheet);
    var pendingSheet = ss.getSheetByName('Đang xử lý');
    
    if (!sheet || !pendingSheet) {
      return { success: false, message: "Không tìm thấy sheet dữ liệu" };
    }
    
    // Tìm dữ liệu theo ID
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var rowIndex = -1;
    var rowData = null;

    for (var i = 1; i < values.length; i++) {
      var rowId = String(values[i][0]).replace(/^['"]|['"]$/g, "").trim();
      var searchId = String(id).replace(/^['"]|['"]$/g, "").trim();
      
      if (rowId === searchId) {
        rowIndex = i + 1; // +1 vì hàng đầu tiên là 1, không phải 0
        rowData = values[i];
        break;
      }
    }
    
    if (rowIndex === -1 || !rowData) {
      return { success: false, message: "Không tìm thấy dữ liệu với ID này trong sheet " + sourceSheet };
    }
    
    // Cập nhật trạng thái thành "Đang xử lý"
    rowData[11] = "Đang xử lý"; // Vị trí của trạng thái
    
    // Thêm vào sheet Đang xử lý
    pendingSheet.appendRow(rowData);
    
    // Xóa từ sheet nguồn
    sheet.deleteRow(rowIndex);
    
    return { success: true, message: "Dữ liệu đã được chuyển về trạng thái Đang xử lý" };
    
  } catch (error) {
    return { success: false, message: "Lỗi khi chuyển dữ liệu: " + error.message };
  }
}

// Hàm lấy danh sách email người phê duyệt
function getApprovers() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User');
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var approvers = [];
    
    // Bỏ qua hàng tiêu đề
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] && data[i][6] === 'Admin' || data[i][6].includes('Quản lý') || data[i][6].includes('Manager')) {
        approvers.push({
          email: data[i][2],
          name: data[i][1],
          role: data[i][6]
        });
      }
    }
    
    return approvers;
  } catch (error) {
    console.error("Lỗi khi lấy danh sách người phê duyệt:", error);
    return [];
  }
}

// Hàm lấy ngày hết hạn phép năm trước từ cài đặt
function getPreviousYearLeaveExpiryDate() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!sheet) return "31/03/" + new Date().getFullYear(); // Giá trị mặc định
    
    // Tìm cài đặt ngày hết hạn từ sheet Config
    // Thêm header nếu chưa có
    if (sheet.getRange("D1").getValue() !== "Ngày hết hạn phép năm trước") {
      sheet.getRange("D1").setValue("Ngày hết hạn phép năm trước");
    }
    
    var expiryDate = sheet.getRange("D2").getValue();
    if (!expiryDate) {
      // Nếu chưa có giá trị, thiết lập giá trị mặc định là 31/03 năm hiện tại
      var defaultDate = new Date(new Date().getFullYear(), 2, 31); // Tháng 3 = 2 trong JS
      sheet.getRange("D2").setValue(defaultDate);
      return Utilities.formatDate(defaultDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
    }
    
    // Nếu giá trị là đối tượng Date, định dạng nó
    if (expiryDate instanceof Date) {
      return Utilities.formatDate(expiryDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
    }
    
    // Nếu giá trị là chuỗi, trả về nguyên chuỗi (giả sử đã đúng định dạng dd/MM/yyyy)
    return expiryDate;
  } catch (error) {
    console.error("Lỗi khi lấy ngày hết hạn phép năm trước:", error);
    return "31/03/" + new Date().getFullYear(); // Giá trị mặc định nếu có lỗi
  }
}

// Hàm cập nhật ngày hết hạn phép năm trước
function updatePreviousYearLeaveExpiryDate(dateStr) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Config');
      sheet.getRange("A1").setValue("Phòng ban");
      sheet.getRange("B1").setValue("Chức vụ");
      sheet.getRange("C1").setValue("Ngày nghỉ lễ");
      sheet.getRange("D1").setValue("Ngày hết hạn phép năm trước");
    }
    
    // Xử lý chuỗi ngày thành đối tượng Date
    var date;
    if (dateStr.includes('/')) {
      var parts = dateStr.split('/');
      if (parts.length === 3) {
        // Chuyển từ dd/MM/yyyy thành Date
        date = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
      } else {
        throw new Error("Định dạng ngày không hợp lệ. Vui lòng sử dụng dd/MM/yyyy.");
      }
    } else {
      date = new Date(dateStr);
    }
    
    if (isNaN(date.getTime())) {
      throw new Error("Ngày không hợp lệ.");
    }
    
    // Cập nhật cài đặt
    sheet.getRange("D2").setValue(date);
    
    return "Đã cập nhật ngày hết hạn phép năm trước thành " + Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
  } catch (error) {
    throw new Error("Lỗi khi cập nhật ngày hết hạn: " + error.message);
  }
}