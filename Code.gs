/**
 * Google Apps Script for Bar Asra Review Tracking
 *
 * SETUP INSTRUCTIONS:
 * 1. Create a new Google Sheet
 * 2. Go to Extensions > Apps Script
 * 3. Delete any default code and paste this entire script
 * 4. Save the script
 * 5. Click the "initializeSheet" function in the dropdown and click Run (play button)
 * 6. Click Deploy > Manage Deployments > Edit > New Version > Deploy
 * 7. Copy your Web App URL
 * 8. Replace WEB_APP_URL in review.html with your Web App URL
 */

// Configuration
const MANAGER_EMAIL = "lbarwe1@gmail.com"; // Update with your email
const ADMIN_PASSWORD = "barasra2024"; // Password for dashboard admin access

/**
 * Run this ONCE to set up your sheets
 */
function initializeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Setup Raw Data sheet (all ratings)
  let rawDataSheet = ss.getSheetByName('Raw Data');
  if (!rawDataSheet) {
    rawDataSheet = ss.insertSheet('Raw Data');
    setupRawDataSheet(rawDataSheet);
  }

  // Setup Feedback sheet (1-4 star feedback)
  let feedbackSheet = ss.getSheetByName('Feedback');
  if (!feedbackSheet) {
    feedbackSheet = ss.insertSheet('Feedback');
    setupFeedbackSheet(feedbackSheet);
  }

  ss.setActiveSheet(rawDataSheet);
  SpreadsheetApp.getUi().alert('✅ Sheets initialized!\n\n📊 Raw Data - All ratings\n💬 Feedback - Customer feedback (1-4 stars)');
}

function setupRawDataSheet(sheet) {
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() !== 'Timestamp') {
    sheet.clear();

    sheet.appendRow([
      'Timestamp',
      'Rating',
      'Source',
      'Date',
      'Time'
    ]);

    // Format header
    const headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1c1a18');
    headerRange.setFontColor('#C7A27A');
    headerRange.setHorizontalAlignment('center');

    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 110);
    sheet.setColumnWidth(5, 90);
    sheet.setFrozenRows(1);
  }
}

function setupFeedbackSheet(sheet) {
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() !== 'Timestamp') {
    sheet.clear();

    sheet.appendRow([
      'Timestamp',
      'Rating',
      'Customer Name',
      'Customer Contact',
      'Feedback',
      'Date'
    ]);

    // Format header
    const headerRange = sheet.getRange(1, 1, 1, 6);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#C7A27A');
    headerRange.setFontColor('#000000');
    headerRange.setHorizontalAlignment('center');

    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 180);
    sheet.setColumnWidth(5, 350);
    sheet.setColumnWidth(6, 110);
    sheet.setFrozenRows(1);
  }
}

/**
 * Handle POST requests from review.html
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.type === 'feedback') {
      return handleFeedback(data, ss);
    } else if (data.type === 'rating') {
      return handleRating(data, ss);
    } else {
      throw new Error('Unknown type: ' + data.type);
    }

  } catch (error) {
    Logger.log('Error: ' + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleRating(data, ss) {
  let rawDataSheet = ss.getSheetByName('Raw Data');
  if (!rawDataSheet) {
    rawDataSheet = ss.insertSheet('Raw Data');
    setupRawDataSheet(rawDataSheet);
  }

  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');

  rawDataSheet.appendRow([
    now,
    data.rating,
    data.source || 'website',
    dateStr,
    timeStr
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleFeedback(data, ss) {
  let feedbackSheet = ss.getSheetByName('Feedback');
  if (!feedbackSheet) {
    feedbackSheet = ss.insertSheet('Feedback');
    setupFeedbackSheet(feedbackSheet);
  }

  let rawDataSheet = ss.getSheetByName('Raw Data');
  if (!rawDataSheet) {
    rawDataSheet = ss.insertSheet('Raw Data');
    setupRawDataSheet(rawDataSheet);
  }

  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // Add to Feedback sheet
  feedbackSheet.appendRow([
    now,
    data.rating,
    data.name || '',
    data.contact || '',
    data.message || '',
    dateStr
  ]);

  // Also add to Raw Data
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
  rawDataSheet.appendRow([
    now,
    data.rating,
    'feedback_form',
    dateStr,
    timeStr
  ]);

  // Send email for feedback (1-4 stars)
  if (data.rating <= 4) {
    sendFeedbackEmail(data);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function sendFeedbackEmail(data) {
  const subject = `📝 New Feedback - ${data.rating} stars`;

  const body = `
A customer has submitted feedback with ${data.rating} stars.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RATING DETAILS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

⭐ Rating: ${data.rating}/5 stars
📅 Date: ${new Date().toLocaleString()}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CUSTOMER INFORMATION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Name: ${data.name || 'Anonymous'}
Contact: ${data.contact || 'Not provided'}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
FEEDBACK
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

${data.message || 'No message'}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

View all feedback: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}
`;

  try {
    MailApp.sendEmail({ to: MANAGER_EMAIL, subject: subject, body: body });
  } catch (error) {
    Logger.log('Email error: ' + error.toString());
  }
}

function doGet(e) {
  const action = e.parameter.action || 'status';

  try {
    let result;

    switch (action) {
      case 'status':
        result = { status: 'running', message: 'Bar Asra Review Tracker is active' };
        break;

      case 'getData':
        // Get all data for dashboard
        result = getDashboardData();
        break;

      case 'updateSettings':
        // Update settings (requires password)
        const password = e.parameter.password;
        if (password !== ADMIN_PASSWORD) {
          result = { success: false, message: 'Invalid password' };
        } else {
          const googleUrl = e.parameter.googleUrl;
          const threshold = e.parameter.threshold;
          const newEmail = e.parameter.email;

          if (googleUrl) updateSetting('googleReviewUrl', googleUrl);
          if (threshold) updateSetting('googleThreshold', parseInt(threshold));
          if (newEmail) updateSetting('managerEmail', newEmail);

          result = { success: true, message: 'Settings updated' };
        }
        break;

      case 'getSettings':
        // Get current settings (limited info)
        result = getSettingsPublic();
        break;

      default:
        result = { status: 'running', message: 'Bar Asra Review Tracker is active' };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawDataSheet = ss.getSheetByName('Raw Data');
  const feedbackSheet = ss.getSheetByName('Feedback');

  const now = new Date();
  const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

  // Get raw data
  const rawData = rawDataSheet ? rawDataSheet.getDataRange().getValues() : [];
  const feedbackData = feedbackSheet ? feedbackSheet.getDataRange().getValues() : [];

  // Calculate stats
  let totalRatings = 0;
  let todayRatings = 0;
  let weekRatings = 0;
  let ratingDistribution = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
  let fiveStarToGoogle = 0;

  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    const rating = parseInt(row[1]) || 0;
    const dateStr = row[3];
    const source = row[2];

    if (rating >= 1 && rating <= 5) {
      totalRatings++;
      ratingDistribution[rating]++;

      if (dateStr === todayStr) {
        todayRatings++;
      }

      const rowDate = new Date(row[0]);
      if (rowDate >= weekAgo) {
        weekRatings++;
      }

      // Source 'website' means they went to Google (5 stars)
      if (source === 'website' && rating === 5) {
        fiveStarToGoogle++;
      }
    }
  }

  // Get recent feedback
  const recentFeedback = [];
  for (let i = feedbackData.length - 1; i >= Math.max(1, feedbackData.length - 20); i--) {
    const row = feedbackData[i];
    recentFeedback.push({
      rating: row[1],
      name: row[2] || 'Anonymous',
      contact: row[3] || '',
      message: row[4] || '',
      date: row[5]
    });
  }

  // Calculate averages
  const avgRating = totalRatings > 0
    ? ((ratingDistribution[1] * 1 + ratingDistribution[2] * 2 + ratingDistribution[3] * 3 + ratingDistribution[4] * 4 + ratingDistribution[5] * 5) / totalRatings).toFixed(2)
    : 0;

  return {
    stats: {
      totalRatings,
      todayRatings,
      weekRatings,
      avgRating,
      fiveStarToGoogle,
      distribution: ratingDistribution
    },
    recentFeedback,
    lastUpdated: now.toISOString()
  };
}

function getSettingsPublic() {
  return {
    googleReviewUrl: 'https://www.google.com/search?q=Bar+Asra',
    googleThreshold: 5,
    // Don't expose email publicly
  };
}

function updateSetting(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = ss.getSheetByName('Settings');

  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('Settings');
    settingsSheet.appendRow(['Setting', 'Value']);
  }

  const data = settingsSheet.getDataRange().getValues();

  // Find and update
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      settingsSheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }

  // Add new
  settingsSheet.appendRow([key, value]);
}
