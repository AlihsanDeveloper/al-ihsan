function doPost(e) {
  try {
    // 1. Parse the incoming JSON data from your website
    var data = JSON.parse(e.postData.contents);

    // 2. VERIFY THE SUPABASE TOKEN (Stops scrapers bypassing the HTML)
    if (!data.accessToken) {
      throw new Error("Unauthorized: No token provided.");
    }

    // FETCH SECRETS FROM SCRIPT PROPERTIES
    var scriptProperties = PropertiesService.getScriptProperties();
    var supabaseUrl = scriptProperties.getProperty('SUPABASE_URL');
    var supabaseAnonKey = scriptProperties.getProperty('SUPABASE_ANON_KEY');

    // Make sure the properties actually exist
    if (!supabaseUrl || !supabaseAnonKey) {
      throw new Error("Server configuration error: Missing Supabase credentials.");
    }

    var verifyUrl = supabaseUrl + "/auth/v1/user";
    var options = {
      method: "get",
      headers: {
        "apikey": supabaseAnonKey,
        "Authorization": "Bearer " + data.accessToken
      },
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(verifyUrl, options);

    // If Supabase rejects the token, stop immediately.
    if (response.getResponseCode() !== 200) {
      throw new Error("Unauthorized: Invalid or expired token.");
    }

    // 3. APPEND TO THE CORRECT SHEET TAB
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(data.membershipType); // "General" or "Privacy"

    if (!sheet) {
      throw new Error("Sheet named '" + data.membershipType + "' not found.");
    }

        // Write the rows based on your exact column structure
    if (data.membershipType === 'General') {
      var lastRow = sheet.getLastRow();

      // 1. Insert a new row right below the last row of the table
      // This automatically copies the formatting of the row above it!
      sheet.insertRowAfter(lastRow);

      // 2. Write the data into the new row
      sheet.getRange(lastRow + 1, 1, 1, 6).setValues([[
        new Date(),            // Timestamp
        data.name,             // Name
        data.email,            // Email
        data.address,          // Address
        data.phone,            // Phone number
        data.terms             // Terms & Conditions
      ]]);

    } else if (data.membershipType === 'Privacy') {
      var lastRow = sheet.getLastRow();

      // Insert new row
      sheet.insertRowAfter(lastRow);

      // Write data (Privacy only has 4 columns)
      sheet.getRange(lastRow + 1, 1, 1, 4).setValues([[
        new Date(),            // Timestamp
        data.name,             // Name (Alias)
        data.email,            // Email
        data.terms             // Terms & Conditions
      ]]);
    }

    // ==========================================
    // ✅ CALL THE EMAIL NOTIFICATION FUNCTION
    // ==========================================
    sendAdminNotification(data);

    // 4. Return success to the website
    return ContentService.createTextOutput(JSON.stringify({"status": "success"}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // Return error to the website
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function sendAdminNotification(data) {
  // 1. Fetch the admin emails dynamically from your Brevo CRM List

  var adminEmails = getAdminEmailsFromBrevo();

  // 2. Fallback: If the list is empty or API fails, use this email so you don't miss alerts
  if (!adminEmails || adminEmails.length === 0) {
    Logger.log("⚠️ No admin emails found in Brevo list. Using fallback email.");
    adminEmails = ["alihsanfoundation219@gmail.com"];
  }

  // 3. Setup the email subject and HTML body
  var subject = "🔔 Al Ihsan - New " + data.membershipType + " Member: " + data.name;

  var htmlBody =
    "<div style='font-family: Arial, sans-serif; background-color: #03032e; padding: 24px; border-radius: 12px; color: #f0f0ea; max-width: 500px;'>" +
      "<h2 style='color: #56aeff; margin-top: 0;'>Al Ihsan Foundation</h2>" +
      "<p style='font-size: 16px; font-weight: bold; color: #f6b758; margin-top: -10px; margin-bottom: 16px;'>New Membership Application</p>" +
      "<p style='color: rgba(240,240,234,0.8);'>A new application has been successfully verified and submitted.</p>" +
      "<table style='width: 100%; border-collapse: collapse; margin-top: 16px; font-size: 14px;'>" +
        "<tr><td style='padding: 8px 0; border-bottom: 1px solid rgba(86,174,255,0.2); color: #888;'>Type:</td><td style='padding: 8px 0; border-bottom: 1px solid rgba(86,174,255,0.2); font-weight: bold; color: #f6b758;'>" + data.membershipType + "</td></tr>" +
        "<tr><td style='padding: 8px 0; border-bottom: 1px solid rgba(86,174,255,0.2); color: #888;'>Name/Alias:</td><td style='padding: 8px 0; border-bottom: 1px solid rgba(86,174,255,0.2); font-weight: bold;'>" + data.name + "</td></tr>" +
        "<tr><td style='padding: 8px 0; border-bottom: 1px solid rgba(86,174,255,0.2); color: #888;'>Email:</td><td style='padding: 8px 0; border-bottom: 1px solid rgba(86,174,255,0.2);'><a href='mailto:" + data.email + "' style='color: #56aeff;'>" + data.email + "</a></td></tr>" +
        (data.address ? "<tr><td style='padding: 8px 0; border-bottom: 1px solid rgba(86,174,255,0.2); color: #888;'>Address:</td><td style='padding: 8px 0; border-bottom: 1px solid rgba(86,174,255,0.2);'>" + data.address + "</td></tr>" : "") +
        (data.phone ? "<tr><td style='padding: 8px 0; border-bottom: 1px solid rgba(86,174,255,0.2); color: #888;'>Phone:</td><td style='padding: 8px 0; border-bottom: 1px solid rgba(86,174,255,0.2);'>" + data.phone + "</td></tr>" : "") +
        "<tr><td style='padding: 8px 0; color: #888;'>T&C Agreed:</td><td style='padding: 8px 0; color: #4ade80;'>" + (data.terms ? "Yes" : "No") + "</td></tr>" +
      "</table>" +
      "<p style='margin-top: 24px; font-size: 12px; color: #666;'>Check the Google Sheet for the complete record.</p>" +
    "</div>";

  // 4. Send individual emails to protect admin privacy
  adminEmails.forEach(function(email) {
    sendViaBrevo(email, "Al Ihsan Admin", subject, htmlBody);
  });
}

/**
 * Fetches contact emails from your specific Brevo CRM list
 */
function getAdminEmailsFromBrevo() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var BREVO_API_KEY = scriptProperties.getProperty('BREVO_API_KEY');
  var BREVO_LIST_ID = scriptProperties.getProperty('BREVO_LIST_ID');

  if (!BREVO_API_KEY || !BREVO_LIST_ID) {
    Logger.log("❌ Missing Brevo API Key or List ID in Script Properties.");
    return [];
  }

  // Fetch up to 500 contacts from the list
  var url = "https://api.brevo.com/v3/contacts/lists/" + BREVO_LIST_ID + "/contacts?limit=500";
  var options = {
    "method": "get",
    "headers": {
      "accept": "application/json",
      "api-key": BREVO_API_KEY
    },
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() === 200) {
      var data = JSON.parse(response.getContentText());
      var emails = [];

      if (data.contacts && data.contacts.length > 0) {
        data.contacts.forEach(function(contact) {
          // Ensure the contact has an email and hasn't unsubscribed/been blacklisted
          if (contact.email && !contact.emailBlacklisted) {
            emails.push(contact.email);
          }
        });
      }

      Logger.log("✅ Found " + emails.length + " valid admin emails from Brevo CRM.");
      return emails;
    } else {
      Logger.log("❌ Failed to fetch contacts. Error: " + response.getContentText());
      return [];
    }
  } catch (e) {
    Logger.log("❌ Apps Script Error: " + e.message);
    return [];
  }
}

/**
 * Sends a transactional email using the Brevo API v3.
 */
function sendViaBrevo(recipientEmail, recipientName, subject, htmlBody) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var BREVO_API_KEY = scriptProperties.getProperty('BREVO_API_KEY');

  if (!BREVO_API_KEY) {
    Logger.log("❌ Brevo API Key not found in Script Properties.");
    return false;
  }

  var SENDER_EMAIL = "alihsanfoundation219@gmail.com";
  var SENDER_NAME = "Al Ihsan";

  var url = "https://api.brevo.com/v3/smtp/email";

  var payload = {
    "sender": { "name": SENDER_NAME, "email": SENDER_EMAIL },
    "to": [ { "email": recipientEmail, "name": recipientName } ],
    "subject": subject,
    "htmlContent": htmlBody
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {
      "accept": "application/json",
      "api-key": BREVO_API_KEY
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();

    if (responseCode === 201) {
      Logger.log("✅ Brevo email sent successfully to: " + recipientEmail);
      return true;
    } else {
      Logger.log("❌ Failed to send via Brevo. Error: " + response.getContentText());
      return false;
    }
  } catch (e) {
    Logger.log("❌ Apps Script Error: " + e.message);
    return false;
  }
}