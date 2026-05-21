const router = require("express").Router()
const getSheets = require("../googleSheet")
const NodeCache = require("node-cache")

// Initialize cache
const cache = new NodeCache({ stdTTL: 300, checkperiod: 120 })
const dataCache = new NodeCache({ stdTTL: 30, checkperiod: 60 })

// Helper functions
function normalizeString(str) {
  if (!str) return ""
  return str.toString().trim().toLowerCase().replace(/\s+/g, " ").replace(/\s*,\s*/g, ",").trim()
}

function parseSpecialParam(paramStr) {
  if (!paramStr || paramStr === 'none' || paramStr === '') return []
  try {
    const decoded = decodeURIComponent(paramStr)
    if (decoded.includes('|||')) {
      return decoded.split('|||').map(p => p.trim())
    }
    return [decoded.trim()]
  } catch (e) {
    if (paramStr.includes('|||')) {
      return paramStr.split('|||').map(p => p.trim())
    }
    return [paramStr.trim()]
  }
}

function isDateInRange(dateStr, startTimestamp, endTimestamp) {
  if (!dateStr) return false
  if (!startTimestamp || !endTimestamp) return true
  
  try {
    let datePart = dateStr.toString().trim().split(' ')[0]
    let plannedDateObj = null
    
    if (datePart.includes('/')) {
      const parts = datePart.split('/')
      plannedDateObj = new Date(parts[2], parts[1] - 1, parts[0])
    } else if (datePart.includes('-')) {
      plannedDateObj = new Date(datePart)
    } else {
      plannedDateObj = new Date(dateStr)
    }
    
    if (isNaN(plannedDateObj.getTime())) return false
    
    plannedDateObj.setHours(0, 0, 0, 0)
    const plannedTimestamp = plannedDateObj.getTime()
    
    return (plannedTimestamp >= startTimestamp && plannedTimestamp <= endTimestamp)
  } catch (e) {
    return false
  }
}

function getDateOnly(dateString) {
  if (!dateString || dateString.toString().trim() === "") return null
  const str = dateString.toString().trim()
  const datePart = str.split(' ')[0]
  
  if (datePart.includes('/')) {
    const parts = datePart.split('/')
    if (parts.length === 3) {
      return `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`
    }
  }
  
  if (datePart.includes('-')) {
    return datePart
  }
  
  try {
    let date = new Date(str)
    if (!isNaN(date.getTime())) {
      return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`
    }
  } catch (e) {
    return null
  }
  return null
}

// ============================================================
// ULTRA FAST: Batch add to Status 3 sheet
// ============================================================
async function batchAddToStatus3Sheet(billsData) {
  try {
    const sheets = await getSheets()
    const email = process.env.CRM_EMAIL || "crm-2@epil.biz"
    const status3SheetName = process.env.STATUS3_SHEET_NAME || "Status 3"
    
    // Get current timestamp once for all rows
    const now = new Date()
    const timestamp = `${now.getDate().toString().padStart(2, '0')}/${(now.getMonth() + 1).toString().padStart(2, '0')}/${now.getFullYear()} ${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}:${now.getSeconds().toString().padStart(2, '0')}`
    
    // Prepare all rows in one go
    const newRows = billsData.map(billData => [
      email,                          // A: Email
      billData.billNo,                // B: Bill Number
      "Step 1",                       // C: ALWAYS "Step 1"
      timestamp,                      // D: Current Timestamp
      billData.plannedDate || "",     // E: Planned Date
      "Next Follow Up date",          // F: Static text
      "",                             // G: Empty
      billData.plannedForLoop || "",  // H: Plan for Loop Date
      "",                             // I: Empty
      "NO"                            // J: Static text "NO"
    ])
    
    // Get current last row
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: `${status3SheetName}!A:A`,
      majorDimension: 'ROWS'
    })
    
    const existingRows = response.data.values || []
    const startRow = existingRows.length + 1
    
    // Prepare batch update data
    const updateData = []
    for (let i = 0; i < newRows.length; i++) {
      updateData.push({
        range: `${status3SheetName}!A${startRow + i}:J${startRow + i}`,
        values: [newRows[i]]
      })
    }
    
    // Execute single batch update for ALL Status 3 entries
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: updateData
      }
    })
    
    console.log(`✅ Batch added ${newRows.length} entries to Status 3`)
    return newRows.length
    
  } catch (error) {
    console.error(`❌ Error batch adding to Status 3:`, error)
    return 0
  }
}

// ============================================================
// CORE FUNCTION: Get pending rows with caching
// ============================================================
async function getPendingRowsWithDateFilter(startTimestamp, endTimestamp, forceRefresh = false) {
  const cacheKey = `pending_rows_${startTimestamp || 'none'}_${endTimestamp || 'none'}`
  
  if (!forceRefresh) {
    const cachedData = dataCache.get(cacheKey)
    if (cachedData) {
      return cachedData
    }
  }
  
  console.log("🔄 Fetching fresh sheet data...")
  const sheets = await getSheets()
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.GOOGLE_SHEET_ID,
    range: `${process.env.SHEET_NAME}!A8:AU`,
    majorDimension: 'ROWS'
  })
  
  const rows = response.data.values || []
  const pendingRows = []
  
  for (let idx = 0; idx < rows.length; idx++) {
    const row = rows[idx]
    const actual3 = row[22]
    const plannedForLoop = row[21]
    
    if ((!actual3 || actual3.toString().trim() === "") && 
        (plannedForLoop && plannedForLoop.toString().trim() !== "")) {
      
      if (startTimestamp && endTimestamp) {
        if (isDateInRange(plannedForLoop, startTimestamp, endTimestamp)) {
          pendingRows.push({ index: idx, row: row })
        }
      } else {
        pendingRows.push({ index: idx, row: row })
      }
    }
  }
  
  const result = {
    rows: pendingRows,
    totalPending: pendingRows.length,
    lastUpdated: Date.now()
  }
  
  dataCache.set(cacheKey, result, 60)
  console.log(`📊 Found ${pendingRows.length} pending rows`)
  return result
}

// ============================================================
// GET ALL PARTIES
// ============================================================
router.get("/parties", async (req, res) => {
  try {
    const { startDate, endDate } = req.query
    
    let startTimestamp = null, endTimestamp = null
    if (startDate && endDate && startDate !== 'undefined' && endDate !== 'undefined' && startDate !== '' && endDate !== '') {
      const startDateObj = new Date(startDate)
      startDateObj.setHours(0, 0, 0, 0)
      startTimestamp = startDateObj.getTime()
      
      const endDateObj = new Date(endDate)
      endDateObj.setHours(23, 59, 59, 999)
      endTimestamp = endDateObj.getTime()
    }
    
    const cacheKey = `parties_${startTimestamp || 'none'}_${endTimestamp || 'none'}`
    const cachedParties = cache.get(cacheKey)
    
    if (cachedParties) {
      return res.json(cachedParties)
    }
    
    const pendingData = await getPendingRowsWithDateFilter(startTimestamp, endTimestamp)
    
    const partiesSet = new Set()
    for (const item of pendingData.rows) {
      const party = item.row[3]
      if (party && party.toString().trim() !== "") {
        partiesSet.add(party.toString().trim())
      }
    }
    
    const parties = Array.from(partiesSet).sort()
    cache.set(cacheKey, parties, 300)
    res.json(parties)
    
  } catch (error) {
    console.error("❌ Error in /parties:", error)
    res.status(500).json({ error: error.message })
  }
})

// ============================================================
// GET CONSIGNEES
// ============================================================
router.get("/consignees", async (req, res) => {
  try {
    const { parties, startDate, endDate } = req.query
    const partyListRaw = parseSpecialParam(parties || '')
    
    let startTimestamp = null, endTimestamp = null
    if (startDate && endDate && startDate !== 'undefined' && endDate !== 'undefined' && startDate !== '' && endDate !== '') {
      const startDateObj = new Date(startDate)
      startDateObj.setHours(0, 0, 0, 0)
      startTimestamp = startDateObj.getTime()
      
      const endDateObj = new Date(endDate)
      endDateObj.setHours(23, 59, 59, 999)
      endTimestamp = endDateObj.getTime()
    }
    
    const cacheKey = `consignees_${parties || 'all'}_${startTimestamp || 'none'}_${endTimestamp || 'none'}`
    const cachedConsignees = cache.get(cacheKey)
    
    if (cachedConsignees) {
      return res.json(cachedConsignees)
    }
    
    const pendingData = await getPendingRowsWithDateFilter(startTimestamp, endTimestamp)
    
    const consigneesMap = new Map()
    const partyListNorm = partyListRaw.map(p => normalizeString(p))
    
    for (const item of pendingData.rows) {
      const party = item.row[3]
      const consignee = item.row[4]
      
      if (!consignee || consignee.toString().trim() === "") continue
      
      if (partyListRaw.length === 0) {
        const original = consignee.toString().trim()
        consigneesMap.set(normalizeString(original), original)
      } else if (party && partyListNorm.includes(normalizeString(party))) {
        const original = consignee.toString().trim()
        consigneesMap.set(normalizeString(original), original)
      }
    }
    
    const uniqueConsignees = Array.from(consigneesMap.values()).sort()
    cache.set(cacheKey, uniqueConsignees, 120)
    res.json(uniqueConsignees)
    
  } catch (error) {
    console.error("❌ Error in /consignees:", error)
    res.status(500).json({ error: error.message })
  }
})

// ============================================================
// GET FILTERED DATA - FAST WITH CACHE
// ============================================================
router.get("/", async (req, res) => {
  try {
    const { startDate, endDate, parties, consignees, skipCache } = req.query
    
    let startTimestamp = null, endTimestamp = null
    if (startDate && endDate && startDate !== 'undefined' && endDate !== 'undefined' && startDate !== '' && endDate !== '') {
      const startDateObj = new Date(startDate)
      startDateObj.setHours(0, 0, 0, 0)
      startTimestamp = startDateObj.getTime()
      
      const endDateObj = new Date(endDate)
      endDateObj.setHours(23, 59, 59, 999)
      endTimestamp = endDateObj.getTime()
    }
    
    const cacheKey = `filtered_${startTimestamp || 'none'}_${endTimestamp || 'none'}_${parties || 'none'}_${consignees || 'none'}`
    
    if (skipCache !== 'true') {
      const cachedResult = cache.get(cacheKey)
      if (cachedResult) {
        return res.json(cachedResult)
      }
    }
    
    const partyListRaw = parseSpecialParam(parties || '')
    const consigneeListRaw = parseSpecialParam(consignees || '')
    const partyListNorm = partyListRaw.map(p => normalizeString(p))
    const consigneeListNorm = consigneeListRaw.map(c => normalizeString(c))
    
    const pendingData = await getPendingRowsWithDateFilter(startTimestamp, endTimestamp, skipCache === 'true')
    
    const filtered = []
    for (const item of pendingData.rows) {
      const row = item.row
      const party = row[3]
      const consignee = row[4]
      
      if (partyListRaw.length > 0 && (!party || !partyListNorm.includes(normalizeString(party)))) continue
      if (consigneeListRaw.length > 0 && (!consignee || !consigneeListNorm.includes(normalizeString(consignee)))) continue
      
      filtered.push({
        billNo: row[2] || "",
        party: party || "",
        consignee: consignee || "",
        billDate: row[1] || "",
        balance: row[6] || "0",
        plannedForLoop: row[21] || "",
        plannedDate: row[20] || "",
        followUp1: row[24] || "",
        BalanceRemaining: row[12] || "0",
        followCount1: row[26] || "0",
        actual3: row[22] || ""
      })
    }
    
    cache.set(cacheKey, filtered, 15)
    res.json(filtered)
    
  } catch (error) {
    console.error("❌ ERROR:", error)
    res.status(500).json({ error: error.message })
  }
})

// ============================================================
// ULTRA FAST BULK FOLLOW-UP UPDATE
// ============================================================
router.post("/update-followup", async (req, res) => {
  try {
    const sheets = await getSheets()
    const { billNumbers, followUpDate } = req.body
    
    if (!billNumbers || !Array.isArray(billNumbers) || billNumbers.length === 0) {
      return res.status(400).json({ error: "Valid bill numbers array is required" })
    }
    
    console.log(`✏️ Bulk update: ${billNumbers.length} bills -> ${followUpDate}`)
    
    // ONE API CALL - Get all data
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: `${process.env.SHEET_NAME}!A8:AU`,
      majorDimension: 'ROWS'
    })
    
    const rows = response.data.values || []
    
    // Create mapping
    const billDataMap = new Map()
    rows.forEach((row, index) => {
      if (row[2]) {
        billDataMap.set(row[2].toString(), {
          index: index,
          row: row,
          billNo: row[2],
          plannedDate: row[20] || "",
          plannedForLoop: row[21] || ""
        })
      }
    })
    
    const formattedDate = getDateOnly(followUpDate) || new Date().toISOString().split('T')[0]
    
    // Prepare FMS update requests
    const fmsUpdateRequests = []
    const successfulBills = []
    const notFound = []
    
    for (const billNumber of billNumbers) {
      const billData = billDataMap.get(billNumber.toString())
      if (billData) {
        const sheetRow = billData.index + 8
        const followCount = parseInt(billData.row[26] || "0")
        
        fmsUpdateRequests.push({
          range: `${process.env.SHEET_NAME}!Y${sheetRow}`,
          values: [[formattedDate]]
        })
        
        fmsUpdateRequests.push({
          range: `${process.env.SHEET_NAME}!AA${sheetRow}`,
          values: [[followCount + 1]]
        })
        
        successfulBills.push(billData)
      } else {
        notFound.push(billNumber)
      }
    }
    
    if (fmsUpdateRequests.length === 0) {
      return res.json({ success: true, updatedCount: 0, notFound, message: "No valid bills found" })
    }
    
    // ONE API CALL - Update all FMS data
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: fmsUpdateRequests
      }
    })
    
    console.log(`✅ FMS updated: ${successfulBills.length} bills`)
    
    // ONE API CALL - Batch add all to Status 3
    const status3Count = await batchAddToStatus3Sheet(successfulBills)
    
    // Clear caches
    cache.flushAll()
    dataCache.flushAll()
    
    console.log(`✅ Complete: ${successfulBills.length} bills, Status 3: ${status3Count} entries`)
    
    res.json({ 
      success: true, 
      updatedCount: successfulBills.length, 
      notFound: notFound,
      status3Entries: status3Count,
      message: `${successfulBills.length} bills updated, ${status3Count} Status 3 entries added`
    })
    
  } catch (err) {
    console.error("❌ Error in bulk update:", err)
    res.status(500).json({ error: err.message })
  }
})

// ============================================================
// FAST SINGLE FOLLOW-UP UPDATE
// ============================================================
router.post("/update-followup-single", async (req, res) => {
  try {
    const sheets = await getSheets()
    const { billNumber, followUpDate } = req.body
    
    if (!billNumber) {
      return res.status(400).json({ error: "Bill number is required" })
    }
    
    console.log(`✏️ Single update: ${billNumber} -> ${followUpDate}`)
    
    // Get fresh data
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: `${process.env.SHEET_NAME}!A8:AU`,
      majorDimension: 'ROWS'
    })
    
    const rows = response.data.values || []
    let rowIndex = -1
    let billData = null
    
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][2] === billNumber) {
        rowIndex = i
        billData = {
          index: i,
          row: rows[i],
          billNo: rows[i][2],
          plannedDate: rows[i][20] || "",
          plannedForLoop: rows[i][21] || ""
        }
        break
      }
    }
    
    if (rowIndex !== -1 && billData) {
      const sheetRow = rowIndex + 8
      const followCount = parseInt(rows[rowIndex][26] || "0")
      const formattedDate = getDateOnly(followUpDate) || new Date().toISOString().split('T')[0]
      
      // Update FMS
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data: [
            {
              range: `${process.env.SHEET_NAME}!Y${sheetRow}`,
              values: [[formattedDate]]
            },
            {
              range: `${process.env.SHEET_NAME}!AA${sheetRow}`,
              values: [[followCount + 1]]
            }
          ]
        }
      })
      
      // Add to Status 3 (using batch function for consistency)
      await batchAddToStatus3Sheet([billData])
      
      cache.flushAll()
      dataCache.flushAll()
      
      res.json({ 
        success: true, 
        message: `Follow-up updated successfully`
      })
    } else {
      res.status(404).json({ success: false, error: "Bill number not found" })
    }
  } catch (err) {
    console.error("❌ Error:", err)
    res.status(500).json({ error: err.message })
  }
})

// ============================================================
// FAST HISTORY
// ============================================================
router.get("/history/:billNumber", async (req, res) => {
  try {
    const { billNumber } = req.params
    const cacheKey = `history_${billNumber}`
    
    const cachedHistory = cache.get(cacheKey)
    if (cachedHistory) {
      return res.json(cachedHistory)
    }
    
    const sheets = await getSheets()
    const status3SheetName = process.env.STATUS3_SHEET_NAME || "Status 3"
    
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: `${status3SheetName}!A:J`,
      majorDimension: 'ROWS'
    })
    
    const rows = response.data.values || []
    const history = []
    
    for (const row of rows) {
      if (row[1] === billNumber) {
        history.push({
          email: row[0],
          billNo: row[1],
          step: row[2],
          timestamp: row[3],
          plannedDate: row[4],
          nextFollowUpText: row[5],
          planForLoopDate: row[7],
          status: row[9]
        })
      }
    }
    
    const result = {
      billNumber: billNumber,
      totalFollowUps: history.length,
      history: history.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))
    }
    
    cache.set(cacheKey, result, 60)
    res.json(result)
    
  } catch (error) {
    console.error("❌ Error fetching history:", error)
    res.status(500).json({ error: error.message })
  }
})

router.post("/clear-cache", (req, res) => {
  cache.flushAll()
  dataCache.flushAll()
  res.json({ success: true })
})

router.get("/health", (req, res) => {
  res.json({ status: "OK", timestamp: new Date().toISOString() })
})

module.exports = router