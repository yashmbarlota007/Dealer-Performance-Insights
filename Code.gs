// ==============================================================================
// YASH MARKETING - DEALER PERFORMANCE & INSIGHTS ENGINE
// Google Apps Script Cloud Engine
// ==============================================================================

/**
 * doGet - Serves the modern Dealer Performance & Insights Web App
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  
  // URL Query Parameters support
  template.empId = (e && e.parameter && e.parameter.empId) ? e.parameter.empId : '';
  template.role = (e && e.parameter && e.parameter.role) ? e.parameter.role : 'Owner';
  template.selectedDealer = (e && e.parameter && e.parameter.dealer) ? e.parameter.dealer : '';
  
  return template.evaluate()
    .setTitle('Yash Marketing - Dealer Performance & Insights Engine')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Include helper for modular HTML/CSS/JS if needed
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Helper to get active spreadsheet or fallback
 */
function getSS_() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    Logger.log("getSS_ Error: " + err.toString());
    return null;
  }
}

/**
 * Helper: Find sheet case-insensitively or return null
 */
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
 * Helper: Parse numbers safely
 */
function parseNum_(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  const cleaned = val.toString().replace(/[^0-9.-]/g, '');
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

/**
 * Helper: Parse date safely
 */
function parseDate_(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Ensures required sheets exist with standard headers
 */
function ensureSystemSheets() {
  const ss = getSS_();
  if (!ss) return { success: false, msg: "Active spreadsheet unavailable" };

  const requiredSheets = [
    {
      name: "Dealer Database",
      headers: ["DealerPhone", "DealerName", "Area", "Beat", "Category", "CreditLimit", "Latitude", "Longitude", "CreatedDate"]
    },
    {
      name: "Order Book",
      headers: ["Date & Time", "OrderID", "ShopName", "DealerPhone", "OrderItems", "TotalAmount", "Payment Mode", "Stage", "DeliveryStatus"]
    },
    {
      name: "Outstanding",
      headers: ["ShopName", "WithBillAmount", "WithoutBillAmount", "OverdueDays", "LastPaymentDate", "CreditStatus"]
    },
    {
      name: "Market Visits",
      headers: ["Timestamp", "StaffName", "ShopName", "Status", "RadialDistance", "Notes", "PhotoURL"]
    },
    {
      name: "Dealer Action Logs",
      headers: ["Timestamp", "ShopName", "ActionType", "Notes", "LoggedBy", "FollowUpDate"]
    }
  ];

  requiredSheets.forEach(function(item) {
    let sheet = ss.getSheetByName(item.name);
    if (!sheet) {
      sheet = ss.insertSheet(item.name);
      sheet.appendRow(item.headers);
      sheet.getRange(1, 1, 1, item.headers.length).setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
    }
  });

  return { success: true, msg: "All system sheets verified." };
}

/**
 * GET DASHBOARD DATA
 * Master aggregator for dealer KPIs, segmentation, and automated insights
 */
function getDealerPerformanceDashboardData(filter) {
  try {
    const ss = getSS_();
    if (!ss) return { success: false, msg: "No active spreadsheet linked." };

    filter = filter || {};
    const selectedArea = filter.area ? filter.area.toString().trim().toLowerCase() : "";
    const selectedBeat = filter.beat ? filter.beat.toString().trim().toLowerCase() : "";
    const selectedTier = filter.tier ? filter.tier.toString().trim().toLowerCase() : "";
    const selectedRange = filter.range || "all"; // "30d", "90d", "this_month", "all"

    const now = new Date();
    const thirtyDaysAgo = new Date(now.getTime() - (30 * 24 * 60 * 60 * 1000));
    const ninetyDaysAgo = new Date(now.getTime() - (90 * 24 * 60 * 60 * 1000));

    // 1. Fetch Dealer Database
    const dealerSheet = findSheetCaseInsensitive_(ss, ["Dealer Database", "Dealers Master", "Dealer Master"]);
    const dealerMap = {};
    const areaSet = new Set();
    const beatSet = new Set();

    if (dealerSheet) {
      const dData = dealerSheet.getDataRange().getValues();
      for (let i = 1; i < dData.length; i++) {
        const row = dData[i];
        const phone = row[0] ? row[0].toString().trim() : "";
        const name = row[1] ? row[1].toString().trim() : "";
        if (!name) continue;

        const area = row[2] ? row[2].toString().trim() : "Unassigned";
        const beat = row[3] ? row[3].toString().trim() : "Default Beat";
        const category = row[4] ? row[4].toString().trim() : "Standard";
        const creditLimit = parseNum_(row[5]) || 50000;
        const lat = row[6] || "";
        const lng = row[7] || "";
        const createdDate = parseDate_(row[8]) || now;

        if (area) areaSet.add(area);
        if (beat) beatSet.add(beat);

        const key = name.toLowerCase();
        dealerMap[key] = {
          phone: phone,
          name: name,
          area: area,
          beat: beat,
          category: category,
          creditLimit: creditLimit,
          latitude: lat,
          longitude: lng,
          createdDate: createdDate,
          totalRevenue: 0,
          totalOrders: 0,
          lastOrderDate: null,
          lastOrderAmount: 0,
          recentOrders: [],
          withBillDue: 0,
          withoutBillDue: 0,
          totalOutstanding: 0,
          overdueDays: 0,
          lastPaymentDate: null,
          totalVisits: 0,
          lastVisitDate: null,
          lastVisitNotes: "",
          monthlySales: {}, // "YYYY-MM" -> amount
          rfmScore: "New",
          healthStatus: "Healthy",
          churnRisk: false
        };
      }
    }

    // 2. Fetch Outstanding Data
    const outSheet = findSheetCaseInsensitive_(ss, ["Outstanding", "Dues", "Outstanding Master"]);
    if (outSheet) {
      const outData = outSheet.getDataRange().getValues();
      for (let i = 1; i < outData.length; i++) {
        const row = outData[i];
        const sName = row[0] ? row[0].toString().trim().toLowerCase() : "";
        if (!sName) continue;

        let target = dealerMap[sName];
        if (!target) {
          // Find matching dealer case-insensitively
          for (let k in dealerMap) {
            if (k === sName || dealerMap[k].name.toLowerCase() === sName) {
              target = dealerMap[k];
              break;
            }
          }
        }

        if (target) {
          const withBill = parseNum_(row[1]);
          const withoutBill = parseNum_(row[2]);
          const overdue = parseNum_(row[3]) || 0;
          const lastPayDate = parseDate_(row[4]);

          target.withBillDue = withBill;
          target.withoutBillDue = withoutBill;
          target.totalOutstanding = withBill + withoutBill;
          target.overdueDays = overdue;
          target.lastPaymentDate = lastPayDate;
        }
      }
    }

    // 3. Fetch Order Book Data
    const orderSheet = findSheetCaseInsensitive_(ss, ["Order Book", "Orders", "Sales"]);
    let systemTotalRevenue = 0;
    let systemTotalOrders = 0;
    const monthlyTrendMap = {}; // "YYYY-MM" -> total
    const productFrequency = {};

    if (orderSheet) {
      const oData = orderSheet.getDataRange().getValues();
      for (let i = 1; i < oData.length; i++) {
        const row = oData[i];
        const orderDate = parseDate_(row[0]);
        const orderId = row[1] ? row[1].toString().trim() : ("ORD-" + i);
        const shopName = row[2] ? row[2].toString().trim() : "";
        const shopPhone = row[3] ? row[3].toString().trim() : "";
        const items = row[4] ? row[4].toString().trim() : "";
        const amount = parseNum_(row[5]);
        const payMode = row[6] ? row[6].toString().trim() : "Credit";
        const stage = row[7] ? row[7].toString().trim() : "Delivered";

        if (!shopName && !shopPhone) continue;

        // Date range filter check
        if (orderDate) {
          if (selectedRange === "30d" && orderDate < thirtyDaysAgo) continue;
          if (selectedRange === "90d" && orderDate < ninetyDaysAgo) continue;
          if (selectedRange === "this_month") {
            if (orderDate.getMonth() !== now.getMonth() || orderDate.getFullYear() !== now.getFullYear()) continue;
          }
        }

        // Aggregate monthly revenue for trend chart
        if (orderDate) {
          const monthKey = orderDate.getFullYear() + "-" + ("0" + (orderDate.getMonth() + 1)).slice(-2);
          monthlyTrendMap[monthKey] = (monthlyTrendMap[monthKey] || 0) + amount;
        }

        // Find dealer
        const dKey = shopName ? shopName.toLowerCase() : "";
        let dObj = dealerMap[dKey];
        if (!dObj) {
          for (let k in dealerMap) {
            if (dealerMap[k].name.toLowerCase() === dKey || (shopPhone && dealerMap[k].phone === shopPhone)) {
              dObj = dealerMap[k];
              break;
            }
          }
        }

        // If dealer not in database, auto-register in map
        if (!dObj && shopName) {
          dObj = {
            phone: shopPhone,
            name: shopName,
            area: "General",
            beat: "Unassigned",
            category: "Standard",
            creditLimit: 50000,
            latitude: "",
            longitude: "",
            createdDate: orderDate || now,
            totalRevenue: 0,
            totalOrders: 0,
            lastOrderDate: null,
            lastOrderAmount: 0,
            recentOrders: [],
            withBillDue: 0,
            withoutBillDue: 0,
            totalOutstanding: 0,
            overdueDays: 0,
            lastPaymentDate: null,
            totalVisits: 0,
            lastVisitDate: null,
            lastVisitNotes: "",
            monthlySales: {},
            rfmScore: "New",
            healthStatus: "Healthy",
            churnRisk: false
          };
          dealerMap[dKey] = dObj;
        }

        if (dObj) {
          dObj.totalRevenue += amount;
          dObj.totalOrders += 1;
          systemTotalRevenue += amount;
          systemTotalOrders += 1;

          if (!dObj.lastOrderDate || (orderDate && orderDate > dObj.lastOrderDate)) {
            dObj.lastOrderDate = orderDate;
            dObj.lastOrderAmount = amount;
          }

          if (orderDate) {
            const mKey = orderDate.getFullYear() + "-" + ("0" + (orderDate.getMonth() + 1)).slice(-2);
            dObj.monthlySales[mKey] = (dObj.monthlySales[mKey] || 0) + amount;
          }

          if (dObj.recentOrders.length < 5) {
            dObj.recentOrders.push({
              orderId: orderId,
              date: orderDate ? orderDate.toISOString().slice(0, 10) : "N/A",
              amount: amount,
              items: items,
              payMode: payMode,
              stage: stage
            });
          }

          // Item affinity frequency
          if (items) {
            const parts = items.split(/[,;\n+]/);
            for (let p = 0; p < parts.length; p++) {
              const cleanP = parts[p].trim();
              if (cleanP.length > 2) {
                productFrequency[cleanP] = (productFrequency[cleanP] || 0) + 1;
              }
            }
          }
        }
      }
    }

    // 4. Fetch Market Visits
    const visitSheet = findSheetCaseInsensitive_(ss, ["Market Visits", "Visit History", "Visits"]);
    if (visitSheet) {
      const vData = visitSheet.getDataRange().getValues();
      for (let i = 1; i < vData.length; i++) {
        const row = vData[i];
        const vDate = parseDate_(row[0]);
        const staff = row[1] ? row[1].toString().trim() : "";
        const shop = row[2] ? row[2].toString().trim().toLowerCase() : "";
        const notes = row[5] ? row[5].toString().trim() : "";

        if (!shop) continue;

        let dObj = dealerMap[shop];
        if (!dObj) {
          for (let k in dealerMap) {
            if (dealerMap[k].name.toLowerCase() === shop) {
              dObj = dealerMap[k];
              break;
            }
          }
        }

        if (dObj) {
          dObj.totalVisits += 1;
          if (!dObj.lastVisitDate || (vDate && vDate > dObj.lastVisitDate)) {
            dObj.lastVisitDate = vDate;
            dObj.lastVisitNotes = (staff ? "[" + staff + "] " : "") + notes;
          }
        }
      }
    }

    // 5. Evaluate Dealer Segmentation, RFM, and Risk Engine
    const dealerList = [];
    let totalOutstandingSum = 0;
    let activeDealersCount = 0;
    let atRiskDealersCount = 0;
    let hibernatingDealersCount = 0;
    let highDebtDealersCount = 0;

    const areaBreakdown = {};
    const tierBreakdown = {
      "Platinum": { count: 0, revenue: 0 },
      "Gold": { count: 0, revenue: 0 },
      "Silver": { count: 0, revenue: 0 },
      "Bronze": { count: 0, revenue: 0 }
    };

    const actionableAlerts = [];

    for (let k in dealerMap) {
      const d = dealerMap[k];

      // Area / Beat / Tier Filters
      if (selectedArea && d.area.toLowerCase() !== selectedArea) continue;
      if (selectedBeat && d.beat.toLowerCase() !== selectedBeat) continue;
      if (selectedTier && d.category.toLowerCase() !== selectedTier) continue;

      totalOutstandingSum += d.totalOutstanding;

      // Area aggregate
      if (!areaBreakdown[d.area]) {
        areaBreakdown[d.area] = { name: d.area, revenue: 0, dealers: 0, outstanding: 0 };
      }
      areaBreakdown[d.area].revenue += d.totalRevenue;
      areaBreakdown[d.area].dealers += 1;
      areaBreakdown[d.area].outstanding += d.totalOutstanding;

      // Days since last order
      let daysSinceLastOrder = 999;
      if (d.lastOrderDate) {
        daysSinceLastOrder = Math.floor((now.getTime() - d.lastOrderDate.getTime()) / (1000 * 60 * 60 * 24));
      }

      // Days since last visit
      let daysSinceLastVisit = 999;
      if (d.lastVisitDate) {
        daysSinceLastVisit = Math.floor((now.getTime() - d.lastVisitDate.getTime()) / (1000 * 60 * 60 * 24));
      }

      // Classification & RFM Engine
      if (d.totalRevenue >= 300000) {
        d.tier = "Platinum";
        tierBreakdown["Platinum"].count++;
        tierBreakdown["Platinum"].revenue += d.totalRevenue;
      } else if (d.totalRevenue >= 100000) {
        d.tier = "Gold";
        tierBreakdown["Gold"].count++;
        tierBreakdown["Gold"].revenue += d.totalRevenue;
      } else if (d.totalRevenue >= 30000) {
        d.tier = "Silver";
        tierBreakdown["Silver"].count++;
        tierBreakdown["Silver"].revenue += d.totalRevenue;
      } else {
        d.tier = "Bronze";
        tierBreakdown["Bronze"].count++;
        tierBreakdown["Bronze"].revenue += d.totalRevenue;
      }

      // Health Status & Churn Risk
      if (daysSinceLastOrder <= 30) {
        d.healthStatus = "Active";
        activeDealersCount++;
      } else if (daysSinceLastOrder <= 75) {
        d.healthStatus = "At Risk";
        d.churnRisk = true;
        atRiskDealersCount++;
      } else {
        d.healthStatus = "Hibernating";
        hibernatingDealersCount++;
      }

      // Debt Risk
      if (d.totalOutstanding > d.creditLimit || d.overdueDays > 45) {
        highDebtDealersCount++;
        d.debtStatus = "Critical";
      } else if (d.totalOutstanding > 0) {
        d.debtStatus = "Normal";
      } else {
        d.debtStatus = "Clear";
      }

      // Push actionable alerts
      if (d.churnRisk && d.totalRevenue > 50000) {
        actionableAlerts.push({
          type: "churn",
          severity: "high",
          title: "Revenue Drop Alert: " + d.name,
          dealer: d.name,
          phone: d.phone,
          message: `High-value dealer (Lifetime: ₹${Math.round(d.totalRevenue).toLocaleString('en-IN')}) inactive for ${daysSinceLastOrder} days.`,
          action: "Schedule Priority Field Visit"
        });
      }

      if (d.totalOutstanding > d.creditLimit) {
        actionableAlerts.push({
          type: "credit",
          severity: "critical",
          title: "Credit Limit Exceeded: " + d.name,
          dealer: d.name,
          phone: d.phone,
          message: `Outstanding ₹${Math.round(d.totalOutstanding).toLocaleString('en-IN')} exceeds limit of ₹${Math.round(d.creditLimit).toLocaleString('en-IN')}.`,
          action: "Hold Pending Orders / Collect Payment"
        });
      } else if (d.overdueDays >= 45 && d.totalOutstanding > 10000) {
        actionableAlerts.push({
          type: "aging",
          severity: "warning",
          title: "Overdue Aging Alert: " + d.name,
          dealer: d.name,
          phone: d.phone,
          message: `Outstanding ₹${Math.round(d.totalOutstanding).toLocaleString('en-IN')} is overdue by ${d.overdueDays} days.`,
          action: "Send WhatsApp Payment Reminder"
        });
      }

      dealerList.push({
        phone: d.phone,
        name: d.name,
        area: d.area,
        beat: d.beat,
        category: d.category,
        tier: d.tier,
        creditLimit: d.creditLimit,
        totalRevenue: d.totalRevenue,
        totalOrders: d.totalOrders,
        aov: d.totalOrders > 0 ? Math.round(d.totalRevenue / d.totalOrders) : 0,
        lastOrderDate: d.lastOrderDate ? d.lastOrderDate.toISOString().slice(0, 10) : "Never",
        daysSinceLastOrder: daysSinceLastOrder === 999 ? "N/A" : daysSinceLastOrder,
        withBillDue: d.withBillDue,
        withoutBillDue: d.withoutBillDue,
        totalOutstanding: d.totalOutstanding,
        overdueDays: d.overdueDays,
        debtStatus: d.debtStatus,
        lastVisitDate: d.lastVisitDate ? d.lastVisitDate.toISOString().slice(0, 10) : "Never",
        daysSinceLastVisit: daysSinceLastVisit === 999 ? "N/A" : daysSinceLastVisit,
        healthStatus: d.healthStatus,
        recentOrders: d.recentOrders
      });
    }

    // Sort Dealers: Top Revenue First
    dealerList.sort(function(a, b) {
      return b.totalRevenue - a.totalRevenue;
    });

    // Top 5 Growth & Top 5 Outstandings
    const topRevenueDealers = dealerList.slice(0, 10);
    const topDebtors = [...dealerList].sort((a, b) => b.totalOutstanding - a.totalOutstanding).slice(0, 10);

    // Limit alerts to top 8 most critical
    const prioritizedAlerts = actionableAlerts.slice(0, 8);

    // Prepare chart trends (Sorted by month key)
    const sortedMonthKeys = Object.keys(monthlyTrendMap).sort();
    const chartLabels = [];
    const chartData = [];
    sortedMonthKeys.forEach(function(k) {
      chartLabels.push(k);
      chartData.push(Math.round(monthlyTrendMap[k]));
    });

    return {
      success: true,
      timestamp: now.toISOString(),
      summary: {
        totalDealers: dealerList.length,
        activeDealers: activeDealersCount,
        atRiskDealers: atRiskDealersCount,
        hibernatingDealers: hibernatingDealersCount,
        highDebtDealers: highDebtDealersCount,
        totalRevenue: Math.round(systemTotalRevenue),
        totalOrders: systemTotalOrders,
        averageOrderValue: systemTotalOrders > 0 ? Math.round(systemTotalRevenue / systemTotalOrders) : 0,
        totalOutstanding: Math.round(totalOutstandingSum)
      },
      chartData: {
        months: chartLabels,
        revenue: chartData
      },
      tierDistribution: tierBreakdown,
      areaBreakdown: Object.values(areaBreakdown),
      topDealers: topRevenueDealers,
      topDebtors: topDebtors,
      dealers: dealerList,
      alerts: prioritizedAlerts,
      filtersList: {
        areas: Array.from(areaSet).sort(),
        beats: Array.from(beatSet).sort(),
        tiers: ["Platinum", "Gold", "Silver", "Bronze"]
      }
    };
  } catch (err) {
    Logger.log("getDealerPerformanceDashboardData Error: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * GET DEALER 360 PROFILE
 * Complete individual 360-degree analytics for modal or deep dive
 */
function getDealer360(dealerName) {
  try {
    const ss = getSS_();
    if (!ss) return { success: false, msg: "Active spreadsheet unavailable" };

    const cleanName = dealerName.trim().toLowerCase();
    const dashboard = getDealerPerformanceDashboardData({ range: "all" });
    if (!dashboard.success) return dashboard;

    let target = null;
    for (let i = 0; i < dashboard.dealers.length; i++) {
      if (dashboard.dealers[i].name.toLowerCase() === cleanName) {
        target = dashboard.dealers[i];
        break;
      }
    }

    if (!target) {
      return { success: false, msg: "Dealer not found: " + dealerName };
    }

    // Fetch Full Orders for this dealer
    const orderSheet = findSheetCaseInsensitive_(ss, ["Order Book", "Orders", "Sales"]);
    const fullOrders = [];
    if (orderSheet) {
      const oData = orderSheet.getDataRange().getValues();
      for (let i = 1; i < oData.length; i++) {
        const row = oData[i];
        const sName = row[2] ? row[2].toString().trim().toLowerCase() : "";
        if (sName === cleanName || sName.includes(cleanName)) {
          const d = parseDate_(row[0]);
          fullOrders.push({
            date: d ? d.toLocaleString('en-IN') : "N/A",
            orderId: row[1] || ("ORD-" + i),
            items: row[4] || "",
            amount: parseNum_(row[5]),
            payMode: row[6] || "Credit",
            stage: row[7] || "Delivered"
          });
        }
      }
    }

    // Fetch Visit Logs
    const visitSheet = findSheetCaseInsensitive_(ss, ["Market Visits", "Visit History", "Visits"]);
    const visitHistory = [];
    if (visitSheet) {
      const vData = visitSheet.getDataRange().getValues();
      for (let i = 1; i < vData.length; i++) {
        const row = vData[i];
        const sName = row[2] ? row[2].toString().trim().toLowerCase() : "";
        if (sName === cleanName || sName.includes(cleanName)) {
          const d = parseDate_(row[0]);
          visitHistory.push({
            date: d ? d.toLocaleString('en-IN') : "N/A",
            staff: row[1] || "Field Staff",
            status: row[3] || "Completed",
            notes: row[5] || ""
          });
        }
      }
    }

    // Fetch Action Logs
    const actionSheet = findSheetCaseInsensitive_(ss, ["Dealer Action Logs", "Action Logs"]);
    const actionLogs = [];
    if (actionSheet) {
      const aData = actionSheet.getDataRange().getValues();
      for (let i = 1; i < aData.length; i++) {
        const row = aData[i];
        const sName = row[1] ? row[1].toString().trim().toLowerCase() : "";
        if (sName === cleanName) {
          const d = parseDate_(row[0]);
          actionLogs.push({
            timestamp: d ? d.toLocaleString('en-IN') : "N/A",
            actionType: row[2] || "",
            notes: row[3] || "",
            loggedBy: row[4] || "",
            followUpDate: row[5] || ""
          });
        }
      }
    }

    return {
      success: true,
      dealer: target,
      orders: fullOrders.reverse().slice(0, 20),
      visits: visitHistory.reverse().slice(0, 15),
      actionLogs: actionLogs.reverse()
    };
  } catch (err) {
    Logger.log("getDealer360 Error: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * LOG DEALER ACTION (Call, WhatsApp, Note, Meeting)
 */
function logDealerAction(payload) {
  try {
    const ss = getSS_();
    if (!ss) return { success: false, msg: "Active spreadsheet unavailable" };

    let sheet = findSheetCaseInsensitive_(ss, ["Dealer Action Logs", "Action Logs"]);
    if (!sheet) {
      sheet = ss.insertSheet("Dealer Action Logs");
      sheet.appendRow(["Timestamp", "ShopName", "ActionType", "Notes", "LoggedBy", "FollowUpDate"]);
      sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
    }

    const timestamp = new Date();
    sheet.appendRow([
      timestamp,
      payload.dealerName || "",
      payload.actionType || "Call",
      payload.notes || "",
      payload.staffName || "Staff",
      payload.followUpDate || ""
    ]);

    return { success: true, msg: "Action successfully recorded!" };
  } catch (err) {
    Logger.log("logDealerAction Error: " + err.toString());
    return { success: false, msg: "Error logging action: " + err.toString() };
  }
}

/**
 * UPDATE DEALER MASTER (Credit Limit, Beat, Area)
 */
function updateDealerMaster(payload) {
  try {
    const ss = getSS_();
    if (!ss) return { success: false, msg: "Active spreadsheet unavailable" };

    const sheet = findSheetCaseInsensitive_(ss, ["Dealer Database", "Dealers Master"]);
    if (!sheet) return { success: false, msg: "Dealer Database sheet not found" };

    const data = sheet.getDataRange().getValues();
    const targetName = (payload.dealerName || "").trim().toLowerCase();
    let targetRow = -1;

    for (let i = 1; i < data.length; i++) {
      const name = data[i][1] ? data[i][1].toString().trim().toLowerCase() : "";
      if (name === targetName) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, msg: "Dealer not found in database: " + payload.dealerName };
    }

    if (payload.area) sheet.getRange(targetRow, 3).setValue(payload.area);
    if (payload.beat) sheet.getRange(targetRow, 4).setValue(payload.beat);
    if (payload.tier) sheet.getRange(targetRow, 5).setValue(payload.tier);
    if (payload.creditLimit !== undefined) sheet.getRange(targetRow, 6).setValue(parseNum_(payload.creditLimit));

    return { success: true, msg: "Dealer details updated successfully!" };
  } catch (err) {
    Logger.log("updateDealerMaster Error: " + err.toString());
    return { success: false, msg: "Error updating dealer: " + err.toString() };
  }
}