// Code.gs

function doGet(e) {
  var page = e.parameter.page || 'index';
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle('BM Apartment Portal');
}

function getSheet(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("Sheet not found: " + sheetName);
  }
  return sheet;
}

function getDashboardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tenantsSheet = getSheet("Tenants");
    var unitsSheet = getSheet("Units");
    var maintenanceSheet = getSheet("Maintenance Requests");
    var rentSheet = getSheet("Rent Payments");

    if (!tenantsSheet || !unitsSheet || !maintenanceSheet || !rentSheet) {
      throw new Error("One or more sheets not found!");
    }

    // 🔹 Fetch Tenants Data
    var tenants = tenantsSheet.getDataRange().getValues();
    var totalTenants = tenants.slice(1).filter(row => String(row[7]).trim() === "Current Tenant").length; // "Status" at index 7

    // 🔹 Fetch Units Data
    var unitsData = unitsSheet.getDataRange().getValues();
    var totalUnits = unitsData.length > 1 ? unitsData.length - 1 : 0;

    // 🔹 Fetch Maintenance Data (Fix Open Requests Count)
    var maintenanceData = maintenanceSheet.getDataRange().getValues();
    var openMaintenanceRequests = maintenanceData.slice(1).filter(row => {
      var status = String(row[5]).trim(); // "Status" column (Index 5)
      return ["New Request", "In Progress", "Pending"].includes(status);
    }).length;

    // 🔹 Calculate Occupied & Vacant Units
    var unitCounts = unitsData.slice(1).reduce((counts, row) => {
      var status = String(row[1]).trim().toLowerCase(); // "Status" column (Index 1)
      counts[status] = (counts[status] || 0) + 1;
      return counts;
    }, {});

    var occupiedUnits = unitCounts["rented"] || 0;
    var vacantUnits = (unitCounts["vacant"] || 0) + (unitCounts["under maintenance vacant"] || 0);

    return {
      success: true,
      totalTenants: totalTenants,
      totalUnits: totalUnits,
      occupiedUnits: occupiedUnits,
      vacantUnits: vacantUnits,
      openMaintenanceRequests: openMaintenanceRequests // Now correctly calculated
    };

  } catch (error) {
    Logger.log("Error in getDashboardData: " + error.message);
    return {
      success: false,
      message: error.message,
      totalTenants: 0,
      totalUnits: 0,
      occupiedUnits: 0,
      vacantUnits: 0,
      openMaintenanceRequests: 0
    };
  }
}


function getAllTenants() {
  try {
    var sheet = getSheet("Tenants");
    if (!sheet) throw new Error('Sheet "Tenants" not found');
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return JSON.stringify({ success: true, tenants: [] });
    var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var tenants = data.map(function (row) {
      var tenantObj = {};
      headers.forEach(function (header, index) {
        tenantObj[header] = row[index] !== undefined ? row[index] : "";
      });
      return tenantObj;
    });
    return JSON.stringify({ success: true, tenants: tenants });
  } catch (error) {
    Logger.log("Error in getAllTenants: " + error.message);
    return JSON.stringify({ success: false, message: error.message });
  }
}

/**
 * Sends a welcome email to the tenant.
 * Uses htmlBody to format the email nicely.
 */
function sendTenantWelcomeEmail(data) {
  var propertyManagerEmail = "bmapartment251@gmail.com";
  var propertyManagerPhone = "+251911249766";
  var apartmentName = "BM Apartment";

  var subject = `Welcome to ${apartmentName} - Your Registration is Complete!`;

  var body = `
    <p>Dear ${data.tenantName},</p>
    <p>Thank you for registering as a tenant at ${apartmentName}. We are excited to welcome you!</p>
    <p>Below are your registration details:</p>
    <ul>
      <li><strong>Unit Number:</strong> ${data.unitNumber}</li>
      <li><strong>Lease Start Date:</strong> ${data.leaseStartDate}</li>
      <li><strong>Lease End Date:</strong> ${data.leaseEndDate}</li>
      <li><strong>Lease Amount:</strong> $${data.leaseAmount}</li>
      <li><strong>Security Deposit Amount:</strong> $${data.securityDeposit}</li>
    </ul>
    <p>If you have any questions, please contact us at ${propertyManagerEmail}.</p>
    <p>Welcome to the community!</p>
    <p>Best regards,<br>Property Management Team</p>
    <hr>
    <p><strong>Contact Us:</strong><br>Email: ${propertyManagerEmail}<br>Phone: ${propertyManagerPhone}</p>
  `;

  try {
    MailApp.sendEmail({
      to: data.email,
      subject: subject,
      htmlBody: body
    });
    return { success: true, message: "Email sent successfully!" };
  } catch (error) {
    Logger.log("Error sending email: " + error.message);
    return { success: false, message: "Error sending email: " + error.message };
  }
}

// Generates a new Request ID based on the last record in the Maintenance Requests sheet.
function getNewRequestID() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Maintenance Requests");
  if (!sheet) {
    throw new Error("Maintenance Requests sheet not found");
  }
  var data = sheet.getDataRange().getValues();
  // Assume header is in row 1 and Request ID is in column 1.
  var lastRequestID = 0;
  for (var i = 1; i < data.length; i++) {
    var reqID = data[i][0];
    if (reqID && reqID.toString().match(/^REQ\d+$/)) {
      var num = parseInt(reqID.replace("REQ", ""), 10);
      if (num > lastRequestID) {
        lastRequestID = num;
      }
    }
  }
  var newNumber = lastRequestID + 1;
  var newID = "REQ" + ("000" + newNumber).slice(-3);
  return newID;
}

// Records a maintenance request in the Maintenance Requests sheet.
// It auto-generates the Request ID and sets the Date Submitted to today.
function recordMaintenanceRequest(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Maintenance Requests");
    if (!sheet) {
      throw new Error("Maintenance Requests sheet not found!");
    }

    // Generate a new Request ID.
    var newRequestID = getNewRequestID();

    // Set Date Submitted to today's date, formatted as "October, 2, 2024".
    var today = new Date();
    var dateSubmitted = Utilities.formatDate(today, Session.getScriptTimeZone(), "MMMM, d, yyyy");

    // Default status for new requests
    var status = "New Request";

    // Append the new request.
    // Column order: Request ID, Unit Number, Tenant Name, Issue Description, Date Submitted, Status
    sheet.appendRow([newRequestID, data.unitNumber, data.tenantName, data.issueDescription, dateSubmitted, status,]);

    return { success: true, message: "Maintenance request recorded successfully with Request ID " + newRequestID };

  } catch (error) {
    Logger.log("Error in recordMaintenanceRequest: " + error.message);
    return { success: false, message: "Error recording maintenance request: " + error.message };
  }
}

function getMaintenanceRequests() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Maintenance Requests");
    if (!sheet) {
      return JSON.stringify({ success: false, message: "Maintenance Requests sheet not found." });
    }

    var data = sheet.getDataRange().getValues();
    var headers = data.shift(); // Remove headers (first row)
    var timeZone = Session.getScriptTimeZone(); // Get timezone

    var requests = data.map(row => {
      let request = {};
      headers.forEach((header, index) => {
        let value = row[index];

        //Check if column contains a date
        if (header === "Date Submitted" || header === "Date Resolved") {
          if (value instanceof Date) {
            value = Utilities.formatDate(value, timeZone, "MMM, dd, yyyy"); // Convert to "Feb, 01, 2025"
          } else {
            value = "N/A"; // Handle invalid dates
          }
        }

        request[header] = value;
      });
      return request;
    });

    return JSON.stringify({ success: true, requests: requests });

  } catch (error) {
    return JSON.stringify({ success: false, message: "Error: " + error.message });
  }
}

function registerTenant(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tenantsSheet = getSheet('Tenants');
    var accessSheet = getSheet('Access List');
    var unitsSheet = getSheet('Units');

    if (!tenantsSheet || !accessSheet || !unitsSheet) {
      throw new Error('One or more sheets not found!');
    }

    var unitsData = unitsSheet.getDataRange().getValues().slice(1);
    var unitIndex = unitsData.findIndex(function (row) { return row[0] == data.unitNumber; });
    if (unitIndex === -1) {
      throw new Error('Error: Unit number not found!');
    }
    if (String(unitsData[unitIndex][1]).trim() !== "Vacant") {
      throw new Error('Error: The selected unit is not vacant!');
    }

    // Append new tenant row.
    // Order: Unit Number, Tenant Name, Phone Number, Email Address, Lease End Date, Security Deposit Amount, Status, Lease Amount, Payment Frequency, First Payment Date.
    tenantsSheet.appendRow([
      data.unitNumber,
      data.tenantName,
      data.phoneNumber,
      data.email,
      data.leaseStartDate,
      data.leaseEndDate,
      data.securityDeposit,   // Security Deposit Amount
      data.status,            // Status
      data.leaseAmount,       // Lease Amount
      data.paymentFrequency   // Payment Frequency integration
    ]);

    // Mark the unit as rented.
    unitsSheet.getRange(unitIndex + 2, 2).setValue("Rented");

    // If tenant is current, add to the Access List.
    if (data.status === "Current Tenant") {
      accessSheet.appendRow([data.email, "Tenant"]);
    }

    // Send welcome email
    sendTenantWelcomeEmail(data);

    return { success: true, message: 'Tenant registered successfully! Unit marked as Rented.' };

  } catch (error) {
    Logger.log("Error in registerTenant: " + error.message);
    return { success: false, message: error.message };
  }
}

function markTenantAsLeft(email, unitNumber) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tenantsSheet = getSheet("Tenants");
    var unitsSheet = getSheet("Units");
    var accessSheet = getSheet("Access List");

    if (!tenantsSheet || !unitsSheet || !accessSheet) {
      throw new Error("One or more sheets not found!");
    }

    var tenantsData = tenantsSheet.getDataRange().getValues();
    var headerRow = tenantsData[0];
    var emailCol = headerRow.indexOf("Email Address");
    var statusCol = headerRow.indexOf("Status");
    var unitNumberCol = headerRow.indexOf("Unit Number");

    if (emailCol === -1 || statusCol === -1 || unitNumberCol === -1) {
      throw new Error("Required headers (Email Address, Status, Unit Number) not found in Tenants sheet");
    }

    var tenantRow = -1;
    for (var i = 1; i < tenantsData.length; i++) {
      if (String(tenantsData[i][emailCol]).trim() === email &&
        String(tenantsData[i][unitNumberCol]).trim() === String(unitNumber).trim()) {
        tenantRow = i + 1;
        tenantsSheet.getRange(tenantRow, statusCol + 1).setValue("Ended Lease");
        tenantsSheet.getRange(tenantRow, emailCol + 1).setValue(String(tenantsData[i][emailCol]).trim() + "_");
        break;
      }
    }

    if (tenantRow === -1) {
      throw new Error("Tenant with the specified email and unit number not found!");
    }

    var unitsData = unitsSheet.getDataRange().getValues();
    var unitFound = false;
    for (var i = 1; i < unitsData.length; i++) {
      if (String(unitsData[i][0]).trim() === String(unitNumber).trim()) {
        unitsSheet.getRange(i + 1, 2).setValue("Vacant");
        unitFound = true;
        break;
      }
    }
    if (!unitFound) {
      Logger.log("Unit " + unitNumber + " not found in Units sheet");
    }

    var accessData = accessSheet.getDataRange().getValues();
    for (var i = accessData.length - 1; i > 0; i--) {
      if (String(accessData[i][0]).trim() === email) {
        accessSheet.deleteRow(i + 1);
        break;
      }
    }

    var dashboardData = getDashboardData();
    dashboardData.message = "Tenant marked as left successfully!";
    return dashboardData;

  } catch (error) {
    Logger.log("Error in markTenantAsLeft: " + error.message);
    return { success: false, message: "Error processing tenant departure: " + error.message };
  }
}

function addMonthsServer(date, months) {
  var d = new Date(date);
  var day = d.getDate();
  d.setMonth(d.getMonth() + months);
  if (d.getDate() < day) {
    d.setDate(0);
  }
  return d;
}
function recordRentPayment(data, fileData, fileName, fileType) {
  try {
    var rentSheet = getSheet("Rent Payments");
    if (!rentSheet) {
      throw new Error("Rent Payments sheet not found!");
    }

    var lastRow = rentSheet.getLastRow();
    var newRentPaymentID = "Rent Pay ID 100"; // Default for first entry

    if (lastRow > 1) { // If there are existing records, get the last Rent Payment ID
      var lastID = rentSheet.getRange(lastRow, 1).getValue();
      var lastNumber = parseInt(lastID.replace("Rent Pay ID ", ""), 10);
      newRentPaymentID = "Rent Pay ID " + (lastNumber + 1);
    }

    // Calculate Next Due Date
    var paymentDateObj = new Date(data.firstPaymentDate);
    var monthsToAdd = {
      "Monthly": 1, "Every 2 Months": 2, "Every 3 Months": 3, "Every 4 Months": 4,
      "Every 6 Months": 6, "Yearly": 12
    }[data.frequency] || 1;
    var nextDueDate = addMonthsServer(paymentDateObj, monthsToAdd);
    var formattedNextDueDate = Utilities.formatDate(nextDueDate, "Africa/Nairobi", "MMMM, d, yyyy");

    // Upload receipt if provided
    var paymentUrl = "";
    if (fileData) {
      var base64Marker = ";base64,";
      var parts = fileData.split(base64Marker);
      if (parts.length === 2) {
        var decoded = Utilities.base64Decode(parts[1]);
        var blob = Utilities.newBlob(decoded, fileType, fileName);
        var folder = DriveApp.getFolderById("1yT4IKjaDXaLSvPcZdups-L75CpoBr5pK");
        var uploadedFile = folder.createFile(blob);
        paymentUrl = uploadedFile.getUrl();
      }
    }

    // Append new record to Rent Payments sheet
    rentSheet.appendRow([
      newRentPaymentID,  //Auto-incremented Rent Payment ID
      data.unitNumber,   // Unit Number
      data.tenantName,   // Tenant Name
      data.frequency,    // Frequency
      data.firstPaymentDate, // Payment Date
      data.amount,       // Amount
      formattedNextDueDate, // Next Due Date
      data.status,       // Status
      data.notes,        // Notes
      paymentUrl         // Payment Receipt URL
    ]);

    return { success: true, message: "Rent payment recorded successfully!" };

  } catch (error) {
    Logger.log("Error in recordRentPayment: " + error.message);
    return { success: false, message: "Error recording rent payment: " + error.message };
  }
}

function getAggregatedRentPayments() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rentSheet = getSheet("Rent Payments");
    if (!rentSheet) {
      throw new Error("Rent Payments sheet not found!");
    }

    var data = rentSheet.getDataRange().getValues();
    if (data.length < 2) return JSON.stringify({ success: true, aggregated: [] });

    // Expected headers: ["Unit Number", "Tenant Name", "Frequency", "Payment Date", "Amount", "Next Due Date", "Status", "Notes", "Payment URL Link"]
    var headers = data[0];
    var unitIndex = headers.indexOf("Unit Number");
    var amountIndex = headers.indexOf("Amount");
    var paymentDateIndex = headers.indexOf("Payment Date");

    var agg = {};
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var unit = row[unitIndex];
      var amount = parseFloat(row[amountIndex]);
      if (isNaN(amount)) amount = 0;
      var paymentDate = row[paymentDateIndex];

      if (!agg[unit]) {
        agg[unit] = {
          unit: unit,
          totalPaid: 0,
          count: 0,
          lastPayment: paymentDate,
          paidMonths: {}
        };
      }

      agg[unit].totalPaid += amount;
      agg[unit].count += 1;

      var newPayment = new Date(paymentDate);
      if (!isNaN(newPayment.getTime())) {
        var currentLast = new Date(agg[unit].lastPayment);
        if (isNaN(currentLast.getTime()) || newPayment > currentLast) {
          agg[unit].lastPayment = paymentDate;
        }
        var monthStr = Utilities.formatDate(newPayment, "Africa/Nairobi", "MMMM");
        agg[unit].paidMonths[monthStr] = true;
      }
    }

    var aggregatedArray = [];
    for (var unit in agg) {
      var record = agg[unit];
      var monthsArray = Object.keys(record.paidMonths);
      record.paidMonths = monthsArray.join(", ");
      aggregatedArray.push(record);
    }

    return JSON.stringify({ success: true, aggregated: aggregatedArray });

  } catch (error) {
    Logger.log("Error in getAggregatedRentPayments: " + error.message);
    return JSON.stringify({ success: false, message: error.message });
  }
}

function getTenantDetails(unitNumber) {
  try {
    var sheet = getSheet("Tenants");
    if (!sheet) throw new Error("Tenants sheet not found!");

    var data = sheet.getDataRange().getValues();
    var headers = data[0]; // Column headers

    // Identify column indexes
    var unitIndex = headers.indexOf("Unit Number");
    var tenantIndex = headers.indexOf("Tenant Name");
    var frequencyIndex = headers.indexOf("Payment Frequency");
    var leaseAmountIndex = headers.indexOf("Lease Amount");
    var statusIndex = headers.indexOf("Status");

    if (unitIndex === -1 || tenantIndex === -1 || frequencyIndex === -1 || leaseAmountIndex === -1 || statusIndex === -1) {
      throw new Error("Required columns not found in Tenants sheet!");
    }

    // Find the tenant with the selected unit and "Current Tenant" status
    for (var i = 1; i < data.length; i++) {
      if (data[i][unitIndex] == unitNumber && data[i][statusIndex] === "Current Tenant") {
        var tenantDetails = {
          tenantName: data[i][tenantIndex],
          paymentFrequency: data[i][frequencyIndex],
          leaseAmount: data[i][leaseAmountIndex]
        };
        Logger.log("Tenant found: " + JSON.stringify(tenantDetails)); // Debugging log
        return tenantDetails;
      }
    }

    Logger.log("No tenant found for Unit Number: " + unitNumber);
    return null; // Return null if no matching tenant found
  } catch (error) {
    Logger.log("Error in getTenantDetails: " + error.message);
    return null;
  }
}

function getNextRentPaymentIDAndTenants() {
  var rentSheet = getSheet("Rent Payments");
  var tenantSheet = getSheet("Tenants");
  if (!rentSheet || !tenantSheet) throw new Error("Sheet not found!");

  // Get next Rent Payment ID
  var rentData = rentSheet.getDataRange().getValues();
  var nextID = (rentData.length < 2) ? "Rent Pay ID 100" :
    "Rent Pay ID " + (parseInt(rentData[rentData.length - 1][0].replace("Rent Pay ID ", "")) + 1);

  // Get "Current Tenant" details from Tenants sheet
  var tenantData = tenantSheet.getDataRange().getValues();
  var headers = tenantData[0];
  var tenants = [];

  tenantData.slice(1).forEach(row => {
    var tenant = {};
    headers.forEach((header, i) => {
      tenant[header] = row[i];
    });

    if (tenant["Status"] === "Current Tenant") {
      tenants.push({
        unitNumber: tenant["Unit Number"],
        tenantName: tenant["Tenant Name"],
        paymentFrequency: tenant["Payment Frequency"],
        leaseAmount: tenant["Lease Amount"]
      });
    }
  });

  return { nextID, tenants };
}


function getRentPayments() {
  try {
    var sheet = getSheet("Rent Payments");
    if (!sheet) throw new Error("Rent Payments sheet not found!");

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return JSON.stringify({ success: true, payments: [] });

    var headers = data[0]; // Get column headers
    var payments = data.slice(1).map(row => {
      var obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] !== undefined ? row[index] : "";
      });
      return obj;
    });

    return JSON.stringify({ success: true, payments: payments });
  } catch (error) {
    Logger.log("Error in getRentPayments: " + error.message);
    return JSON.stringify({ success: false, message: error.message });
  }
}

function getVacantUnits() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var unitsSheet = getSheet('Units');
  if (!unitsSheet) {
    return [];
  }
  var unitsData = unitsSheet.getDataRange().getValues();
  // Assumes column B (index 1) contains the unit status and column A (index 0) the unit number.
  return unitsData.slice(1)
    .filter(function (row) {
      return String(row[1]).trim() === "Vacant";
    })
    .map(function (row) {
      return row[0];
    });
}
function getNewBuildingMaintenanceID() {
  var sheet = getSheet("Building Maintenance");
  if (!sheet) {
    throw new Error("Building Maintenance sheet not found!");
  }
  var data = sheet.getDataRange().getValues();

  var lastID = 0;
  for (var i = 1; i < data.length; i++) {
    var reqID = data[i][0];
    if (reqID && reqID.toString().match(/^Main Req ID \d+$/)) {
      var num = parseInt(reqID.replace("Main Req ID ", ""), 10);
      if (num > lastID) {
        lastID = num;
      }
    }
  }
  var newID = "Main Req ID " + ("000" + (lastID + 1)).slice(-3);
  return newID;
}

function recordBuildingMaintenance(data) {
  try {
    var sheet = getSheet("Building Maintenance");
    if (!sheet) throw new Error("Building Maintenance sheet not found!");

    var newRequestID = getNewBuildingMaintenanceID();

    // Convert date strings to Google Sheets date format
    var dueDate = data.dueDate ? new Date(data.dueDate) : "";
    var dateCompleted = data.dateCompleted ? new Date(data.dateCompleted) : "";

    sheet.appendRow([
      newRequestID,
      data.taskName,
      dueDate, // Ensure dates are stored properly
      data.completionStatus,
      data.expenseAmount,
      dateCompleted, // Ensure dates are stored properly
      data.notes
    ]);

    return { success: true, message: "Building Maintenance recorded successfully!" };

  } catch (error) {
    Logger.log("Error in recordBuildingMaintenance: " + error.message);
    return { success: false, message: error.message };
  }
}
function getBuildingMaintenanceRecords() {
  try {
    var sheet = getSheet("Building Maintenance");
    if (!sheet) throw new Error("Building Maintenance sheet not found!");

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return JSON.stringify({ success: true, maintenanceRecords: [] });

    var headers = data[0]; // Column headers
    var maintenanceRecords = data.slice(1).map(row => {
      var obj = {};
      headers.forEach((header, index) => {
        var value = row[index];

        // Automatically format date values as "Feb, 01, 2025"
        if (header.includes("Date") && value instanceof Date) {
          obj[header] = formatDateDisplay(value);
        } else {
          obj[header] = value !== undefined ? value : "";
        }
      });
      return obj;
    });

    return JSON.stringify({ success: true, maintenanceRecords });
  } catch (error) {
    Logger.log("Error in getBuildingMaintenanceRecords: " + error.message);
    return JSON.stringify({ success: false, message: error.message });
  }
}

// 🔹 Helper function to format dates as "Feb, 01, 2025"
function formatDateDisplay(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return "N/A";

  return date.toLocaleDateString("en-US", { month: "short", day: "2-digit", year: "numeric" }).replace(" ", ", ");
}

function deleteRentPayment(data) {
  try {
    var sheet = getSheet("Rent Payments");
    if (!sheet) throw new Error("Rent Payments sheet not found!");

    var dataRange = sheet.getDataRange().getValues();
    var rentPayID = data["Rent Payment ID"]; // Make sure you're passing this from the frontend

    for (var i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] === rentPayID) {  // Match Rent Payment ID (Column 1)
        sheet.deleteRow(i + 1);
        return { success: true, message: "Rent payment deleted successfully!" };
      }
    }

    return { success: false, message: "Record not found in Rent Payments." };

  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deleteMaintenanceRequest(data) {
  try {
    var sheet = getSheet("Maintenance Requests");
    if (!sheet) throw new Error("Maintenance Requests sheet not found!");

    var dataRange = sheet.getDataRange().getValues();
    for (var i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] == data["Request ID"]) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Maintenance request deleted successfully!" };
      }
    }
    return { success: false, message: "Record not found in Maintenance Requests." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}
function deleteBuildingMaintenance(data) {
  try {
    var sheet = getSheet("Building Maintenance");
    if (!sheet) throw new Error("Building Maintenance sheet not found!");

    var dataRange = sheet.getDataRange().getValues();
    for (var i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] == data["Maintenance REQ ID"]) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Building maintenance record deleted successfully!" };
      }
    }
    return { success: false, message: "Record not found in Building Maintenance." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}
function updateRentPayment(updatedData, fileData, fileName, fileType) {
  try {
    var sheet = getSheet("Rent Payments");
    if (!sheet) throw new Error("Rent Payments sheet not found!");

    var dataRange = sheet.getDataRange().getValues();
    var newURL = updatedData["Payment URL Link"];

    // If a new file is uploaded, save it in Google Drive
    if (fileData) {
      var base64Marker = ";base64,";
      var parts = fileData.split(base64Marker);
      if (parts.length === 2) {
        var decoded = Utilities.base64Decode(parts[1]);
        var blob = Utilities.newBlob(decoded, fileType, fileName);
        var folder = DriveApp.getFolderById("1yT4IKjaDXaLSvPcZdups-L75CpoBr5pK");  // Replace with your Drive folder ID
        var uploadedFile = folder.createFile(blob);
        newURL = uploadedFile.getUrl();
      }
    }

    // Find the row by Rent Payment ID
    for (var i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] == updatedData["Rent Payment ID"]) {  //Use Rent Payment ID
        updatedData["Payment URL Link"] = newURL;

        // Set values explicitly in the correct column order
        sheet.getRange(i + 1, 1, 1, 10).setValues([[  //Now includes 10 columns
          updatedData["Rent Payment ID"],  // Column 1
          updatedData["Unit Number"],      // Column 2
          updatedData["Tenant Name"],      // Column 3
          updatedData["Frequency"],        // Column 4
          updatedData["Payment Date"],     // Column 5
          updatedData["Amount"],           // Column 6
          updatedData["Next Due Date"],    // Column 7
          updatedData["Status"],           // Column 8
          updatedData["Notes"],            // Column 9
          updatedData["Payment URL Link"]  // Column 10
        ]]);

        return { success: true, message: "Rent payment updated successfully!" };
      }
    }

    return { success: false, message: "Record not found." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}


function updateMaintenanceRequest(oldData, updatedData) {
  try {
    var sheet = getSheet("Maintenance Requests");
    if (!sheet) throw new Error("Maintenance Requests sheet not found!");

    var dataRange = sheet.getDataRange().getValues();

    for (var i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] == oldData["Request ID"]) {
        sheet.getRange(i + 1, 1, 1, 9).setValues([[
          updatedData["Request ID"],
          updatedData["Unit Number"],
          updatedData["Tenant Name"],
          updatedData["Issue Description"],
          updatedData["Date Submitted"],
          updatedData["Status"],
          updatedData["Date Resolved"],
          updatedData["Days to Complete"],
          updatedData["Notes"]
        ]]);

        return { success: true, message: "Maintenance request updated successfully!" };
      }
    }
    return { success: false, message: "Record not found." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function updateBuildingMaintenance(oldData, updatedData) {
  try {
    var sheet = getSheet("Building Maintenance");
    if (!sheet) throw new Error("Building Maintenance sheet not found!");

    var dataRange = sheet.getDataRange().getValues();

    for (var i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] == oldData["Maintenance REQ ID"]) {
        sheet.getRange(i + 1, 1, 1, 7).setValues([[
          updatedData["Maintenance REQ ID"],
          updatedData["Task Name"],
          updatedData["Due Date"],
          updatedData["Completion Status"],
          updatedData["Expense Amount"],
          updatedData["Date Completed"],
          updatedData["Notes"]
        ]]);

        return { success: true, message: "Building maintenance record updated successfully!" };
      }
    }
    return { success: false, message: "Record not found." };
  } catch (error) {
    return { success: false, message: error.message };
  }
}
function getRentDueRecords() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rent Payments");
    if (!sheet) {
      Logger.log("Rent Payments sheet not found!");
      return [];
    }

    var data = sheet.getDataRange().getValues();
    Logger.log("Raw Rent Due Data: " + JSON.stringify(data));

    if (data.length <= 1) {
      Logger.log("No Rent Due records found.");
      return [];
    }

    var headers = data[0]; // Get the headers
    Logger.log("Column Headers: " + JSON.stringify(headers));

    var rentDueRecords = {};
    var today = new Date();

    for (var i = 1; i < data.length; i++) {
      var record = {};
      for (var j = 0; j < headers.length; j++) {
        var key = headers[j].trim();
        var value = data[i][j];

        // Handle Date Conversion Correctly
        if (key.includes("Date")) {
          if (value instanceof Date) {
            value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else if (typeof value === "number") {
            value = Utilities.formatDate(new Date(value), Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else if (typeof value === "string") {
            value = value.trim();
          }
        }

        record[key] = value;
      }

      var unitNumber = record["Unit Number"];
      var nextRentDateStr = record["Next Due Date"]; // FIXED COLUMN NAME
      var amountDue = parseFloat(record["Amount"]) || 0;

      Logger.log(`🔹 Checking: ${unitNumber} | Next Due Date: ${nextRentDateStr} | Amount: ${amountDue}`);

      if (!unitNumber || !nextRentDateStr) {
        Logger.log("Skipping record: Missing Unit Number or Next Due Date");
        continue;
      }

      // Convert date string to Date object
      var dueDate = new Date(nextRentDateStr);
      if (isNaN(dueDate)) {
        Logger.log("Skipping invalid date: " + nextRentDateStr);
        continue;
      }

      // Only consider rent due dates that are today or in the future
      if (dueDate < today) {
        Logger.log("Skipping past due date: " + nextRentDateStr);
        continue;
      }

      // Store only the next upcoming rent due date & amount per unit
      if (!rentDueRecords[unitNumber] || dueDate < new Date(rentDueRecords[unitNumber]["Next Due Date"])) {
        rentDueRecords[unitNumber] = {
          "Unit Number": unitNumber,
          "Tenant Name": record["Tenant Name"],
          "Next Due Date": nextRentDateStr,
          "Amount Due": amountDue
        };
      }
    }

    var finalData = Object.values(rentDueRecords);
    Logger.log("Final Rent Due Data: " + JSON.stringify(finalData));

    return finalData;

  } catch (error) {
    Logger.log("Error in getRentDueRecords(): " + error.toString());
    return [];
  }
}

function checkUserAccess() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Access List");
    if (!sheet) {
        Logger.log("Access List sheet not found!");
        return { access: false, message: "Access List sheet not found!" };
    }

    var data = sheet.getDataRange().getValues();
    var userEmail = Session.getEffectiveUser().getEmail();
    
    // Debugging logs
    if (!userEmail) {
        Logger.log("Failed to get user email. Possible permission issue.");
        return { access: false, message: "Error retrieving your email. Ensure you are logged in with the correct account." };
    }

    userEmail = userEmail.trim().toLowerCase(); // Normalize email format
    Logger.log("Checking access for: " + userEmail);

    for (var i = 1; i < data.length; i++) {
        var email = (data[i][0] || "").trim().toLowerCase(); // Normalize stored emails
        var role = (data[i][1] || "").trim().toLowerCase();  // Normalize stored roles

        Logger.log("Checking email:", email, "| Role:", role);

        if (email === userEmail) {
            if (role === "property manager" || role === "tenant") {
                Logger.log("Access granted to:", userEmail);
                return { access: true, email: userEmail };
            } else {
                Logger.log("Access denied: Invalid role", role);
                return { access: false, message: "Access Denied: Only Property Managers can access this application." };
            }
        }
    }

    Logger.log("Access denied: Email not found");
    return { access: false, message: "Access Denied: Your email is not registered." };
}

function getNextRentPaymentID() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rent Payments");
    if (!sheet) {
        return "Rent Pay ID 100";
    }

    var data = sheet.getDataRange().getValues();
    var nextID = (data.length < 2) ? "Rent Pay ID 100" :
        "Rent Pay ID " + (parseInt(data[data.length - 1][0].replace("Rent Pay ID ", ""), 10) + 1);

    return nextID;
}


function recordRentPaymentFromRecordModal(data, fileData, fileName, fileType) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rent Payments");
    if (!sheet) {
        return { success: false, message: "Rent Payments sheet not found!" };
    }

    // Get last Rent Pay ID and generate the next one
    var dataRange = sheet.getDataRange().getValues();
    var nextID = (dataRange.length < 2) ? "Rent Pay ID 100" :
        "Rent Pay ID " + (parseInt(dataRange[dataRange.length - 1][0].replace("Rent Pay ID ", ""), 10) + 1);

    Logger.log("Next Rent Payment ID: " + nextID);

    // Prepare row data
    var newRow = [
        nextID,  // Auto-generated Rent Payment ID (Corrected)
        data["Unit Number"],
        data["Tenant Name"],
        data["Frequency"],
        data["Payment Date"],
        data["Amount"],
        data["Next Due Date"],
        data["Status"],
        data["Notes"]
    ];

    // Handle file upload if provided
    if (fileData) {
        var folder = DriveApp.getFolderById("1yT4IKjaDXaLSvPcZdups-L75CpoBr5pK"); // Update with your Drive folder ID
        var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(fileData.split(",")[1]), fileType, fileName));
        newRow.push(file.getUrl());
    } else {
        newRow.push(""); // No receipt uploaded
    }

    // Append row to the sheet
    sheet.appendRow(newRow);
    Logger.log("Rent Payment recorded successfully with ID: " + nextID);

    return { success: true, message: "Rent Payment recorded successfully!" };
}
