// ==============================================================================
// YASH MARKETING - DEALER PERFORMANCE & INSIGHTS ENGINE
// Google Apps Script Cloud Engine + Live Dealer Activity & Product Spy Telemetry
// ==============================================================================

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  template.empId = (e && e.parameter && e.parameter.empId) ? e.parameter.empId : '';
  template.role = (e && e.parameter && e.parameter.role) ? e.parameter.role : 'Owner';
  template.selectedDealer = (e && e.parameter && e.parameter.dealer) ? e.parameter.dealer : '';
  
  return template.evaluate()
    .setTitle('Yash Marketing - Dealer Performance, Opens & Product Views Spy Engine')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSS_() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    return null;
  }
}

function findSheetCaseInsensitive_(ss, nameAliases) {
  if (!ss) return null;
  const sheets = ss.getSheets();
  for (let s = 0; s < sheets.length; s++) {
    const sName = sheets[s].getName().trim().toLowerCase();
    for (let a = 0; a < nameAliases.length; a++) {
      if (sName === nameAliases[a].trim().toLowerCase()) {
        return sheets[s];
      }
    }
  }
  return null;
}

/**
 * GET DEALER ACTIVITY TELEMETRY
 * Reads from "Activity Logs" sheet populated by Yash-Marketing-App logEvent
 * Tracks daily app open counts and exact products viewed per dealer
 */
function getDealerActivityTelemetry() {
  try {
    const ss = getSS_();
    if (!ss) return { success: false, msg: "Active spreadsheet unavailable" };

    const logSheet = findSheetCaseInsensitive_(ss, ["Activity Logs", "ActivityLogs", "Dealer Activity"]);
    if (!logSheet) {
      return { success: true, opensToday: {}, productViews: {}, recentLogs: [] };
    }

    const data = logSheet.getDataRange().getValues();
    const today = new Date().toISOString().slice(0, 10);
    const opensToday = {};
    const productViews = {}; // dealer -> [ { product, count, lastTime } ]
    const recentLogs = [];

    // Columns: [Timestamp, Phone, ShopName, Event, Details]
    for (let i = data.length - 1; i >= 1 && recentLogs.length < 100; i--) {
      const row = data[i];
      const d = row[0] ? new Date(row[0]) : null;
      const phone = row[1] ? row[1].toString().trim() : "";
      const shop = row[2] ? row[2].toString().trim() : "Unknown";
      const event = row[3] ? row[3].toString().trim() : "";
      const details = row[4] ? row[4].toString().trim() : "";

      const dStr = d && !isNaN(d.getTime()) ? d.toISOString().slice(0, 10) : "";

      if (event === "App Opened" && dStr === today) {
        opensToday[shop] = (opensToday[shop] || 0) + 1;
      }

      if (event === "Viewed Product") {
        if (!productViews[shop]) productViews[shop] = {};
        productViews[shop][details] = (productViews[shop][details] || 0) + 1;
      }

      recentLogs.push({
        timestamp: d ? d.toLocaleString('en-IN') : "N/A",
        phone: phone,
        dealer: shop,
        event: event,
        details: details
      });
    }

    return {
      success: true,
      opensToday: opensToday,
      productViews: productViews,
      recentLogs: recentLogs
    };
  } catch (err) {
    Logger.log("getDealerActivityTelemetry Error: " + err.toString());
    return { success: false, error: err.toString() };
  }
}
