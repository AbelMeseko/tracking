// ============================================
// CONFIGURATION - GOOGLE SHEET
// ============================================
const SHEET_ID = '1dB0uCnReB_y7Ocu_hgXfRbQsjJtm16azQhnUpb2SWw0';

// Tab GIDs - CORRECTED with your actual GIDs from the links
const TABS = {
    'MAIN': { gid: '0', type: 'main' },
    'BD78NGZN': { gid: '560267061', type: 'vehicle' },
    'CS44GHNZ': { gid: '1291826327', type: 'vehicle' },
    'DG28ZLZN': { gid: '1216996054', type: 'vehicle' }
};

// Vehicle list
const VEHICLES = ['BD78NGZN', 'CS44GHNZ', 'DG28ZLZN'];

// State
let allSheetData = {};
let currentView = 'dashboard';
let distanceChart = null;
let dataLoaded = false;

// Daily totals by source and vehicle
let dailyTotals = {
    MAIN: {},
    VEHICLE: {}
};

// Store all vehicle tab records for matching
let vehicleRecords = {
    'BD78NGZN': [],
    'CS44GHNZ': [],
    'DG28ZLZN': []
};

// Current date range (for dashboard)
let currentDateRange = {
    startDate: null,
    endDate: null
};

// Filter state for MAIN tab
let mainFilters = {
    driver: 'all',
    truck: 'all',
    startDate: '',
    endDate: ''
};

// Filter state for vehicle tabs
let vehicleTabFilters = {
    startDate: '',
    endDate: ''
};

// Safe element getter
function getElement(id) {
    return document.getElementById(id);
}

// ============================================
// ROBUST DATE NORMALIZATION - Handles all your formats
// ============================================
function normalizeDate(dateStr, sourceTab = 'unknown') {
    if (!dateStr) return '';
    
    try {
        // Clean the string
        let cleaned = dateStr.toString().trim();
        
        // Handle empty or invalid
        if (cleaned === '' || cleaned === '0') return '';
        
        // Check if it's already in YYYY-MM-DD format (what we want to output)
        if (cleaned.match(/^\d{4}-\d{2}-\d{2}/)) {
            return cleaned.substring(0, 10);
        }
        
        // Extract date part (remove time if present)
        let datePart = cleaned.split(' ')[0];
        
        // Handle different separators
        let parts = datePart.split(/[\/\-\.]/);
        
        if (parts.length === 3) {
            let part1 = parts[0].trim();
            let part2 = parts[1].trim();
            let part3 = parts[2].trim();
            
            // Check if it's YYYY/MM/DD format (like 2026/02/03)
            if (part1.length === 4 && parseInt(part1) > 2000) {
                // It's YYYY/MM/DD
                let year = part1;
                let month = part2.padStart(2, '0');
                let day = part3.padStart(2, '0');
                return `${year}-${month}-${day}`;
            }
            
            // Handle 2-digit years
            if (part3.length === 2) {
                if (parseInt(part3) >= 24 && parseInt(part3) <= 99) {
                    part3 = '19' + part3;
                } else {
                    part3 = '20' + part3;
                }
            }
            
            // Handle 0026 format
            if (part3 === '0026') {
                part3 = '2026';
            }
            
            // Handle cases where year might be in the middle (MM/YYYY/DD)
            if (part2.length === 4 && parseInt(part2) > 2000) {
                // Format is MM/YYYY/DD
                let month = part1.padStart(2, '0');
                let year = part2;
                let day = part3.padStart(2, '0');
                return `${year}-${month}-${day}`;
            }
            
            // Determine format based on tab and values
            let day, month, year;
            
            // For MAIN tab, assume MM/DD/YYYY (American format)
            if (sourceTab === 'MAIN') {
                // MAIN tab uses MM/DD/YYYY
                month = part1.padStart(2, '0');
                day = part2.padStart(2, '0');
                year = part3;
                
                // Validate month (should be 1-12)
                let monthNum = parseInt(month);
                if (monthNum < 1 || monthNum > 12) {
                    // If month is invalid, try swapping
                    let temp = day;
                    day = month;
                    month = temp;
                }
            } 
            // For vehicle tabs (BD78NGZN, CS44GHNZ, DG28ZLZN), assume DD/MM/YYYY (South African format)
            else {
                // Vehicle tabs use DD/MM/YYYY
                day = part1.padStart(2, '0');
                month = part2.padStart(2, '0');
                year = part3;
                
                // Validate day (should be 1-31)
                let dayNum = parseInt(day);
                if (dayNum < 1 || dayNum > 31) {
                    // If day is invalid, try swapping
                    let temp = day;
                    day = month;
                    month = temp;
                }
            }
            
            // Validate month (should be 1-12)
            let monthNum = parseInt(month);
            if (monthNum < 1 || monthNum > 12) {
                // If month is still invalid, try to detect from values
                let part1Num = parseInt(part1);
                let part2Num = parseInt(part2);
                
                if (part1Num >= 1 && part1Num <= 12 && part2Num >= 1 && part2Num <= 31) {
                    // part1 is likely month
                    month = part1.padStart(2, '0');
                    day = part2.padStart(2, '0');
                } else if (part2Num >= 1 && part2Num <= 12 && part1Num >= 1 && part1Num <= 31) {
                    // part2 is likely month
                    month = part2.padStart(2, '0');
                    day = part1.padStart(2, '0');
                } else {
                    // Default to safe values
                    month = '01';
                    day = '01';
                }
            }
            
            // Validate day (should be 1-31)
            let dayNum = parseInt(day);
            if (dayNum < 1 || dayNum > 31) {
                day = '01';
            }
            
            return `${year}-${month}-${day}`;
        }
        
        // If no slashes, try parsing as ISO
        let date = new Date(cleaned);
        if (!isNaN(date.getTime())) {
            let year = date.getFullYear();
            let month = String(date.getMonth() + 1).padStart(2, '0');
            let day = String(date.getDate()).padStart(2, '0');
            return `${year}-${month}-${day}`;
        }
        
        return cleaned;
        
    } catch (e) {
        console.warn('Date normalization error:', e, 'for date:', dateStr, 'in tab:', sourceTab);
        return '';
    }
}

// Test function to verify date parsing
function testDateNormalization() {
    console.log("Testing date normalization:");
    const testDates = [
        { date: "2/25/2026", tab: "MAIN" },        // Should be 2026-02-25 (MM/DD)
        { date: "02/13/2026", tab: "MAIN" },       // Should be 2026-02-13 (MM/DD)
        { date: "01/02/2026", tab: "BD78NGZN" },   // Should be 2026-01-02 (DD/MM)
        { date: "3/2/2026", tab: "BD78NGZN" },     // Should be 2026-03-02 (DD/MM)
        { date: "2026/02/03", tab: "BD78NGZN" },   // Should be 2026-02-03 (YYYY/MM/DD)
        { date: "08/02/2026 7:15:34", tab: "BD78NGZN" }, // Should be 2026-08-02
        { date: "3/4/2026", tab: "MAIN" }          // Should be 2026-03-04 (MM/DD)
    ];
    
    testDates.forEach(item => {
        console.log(`${item.date} (${item.tab}) -> ${normalizeDate(item.date, item.tab)}`);
    });
}

function isValidHeader(header) {
    if (!header || typeof header !== 'string') return false;
    header = header.trim();
    if (header === '') return false;
    const invalid = ['column', 'null', 'undefined', 'nan', '-', ''];
    return !invalid.includes(header.toLowerCase());
}

function isColumnNumberHeader(header) {
    if (!header || typeof header !== 'string') return false;
    const pattern = /^\s*[Cc][Oo][Ll][Uu][Mm][Nn]\s*\d+\s*$/;
    return pattern.test(header);
}

function findColumnIndex(headers, searchTerm) {
    if (!headers || !Array.isArray(headers)) return -1;
    
    // Try exact match first
    let idx = headers.findIndex(header => 
        header && typeof header === 'string' && header.toLowerCase().trim() === searchTerm.toLowerCase().trim()
    );
    
    if (idx !== -1) return idx;
    
    // Try includes match
    idx = headers.findIndex(header => 
        header && typeof header === 'string' && header.toLowerCase().includes(searchTerm.toLowerCase())
    );
    
    return idx;
}

function extractRegistration(regStr) {
    if (!regStr) return '';
    if (regStr.includes('BD78')) return 'BD78NGZN';
    if (regStr.includes('CS44')) return 'CS44GHNZ';
    if (regStr.includes('DG28')) return 'DG28ZLZN';
    return regStr;
}

// Get last 7 days dates in YYYY-MM-DD format
function getLast7Days() {
    const dates = [];
    const today = new Date();
    
    for (let i = 6; i >= 0; i--) {
        const date = new Date(today);
        date.setDate(today.getDate() - i);
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        dates.push(`${year}-${month}-${day}`);
    }
    
    return dates;
}

// Set date range to last 7 days
function setLast7Days() {
    const last7Days = getLast7Days();
    const startDate = last7Days[0];
    const endDate = last7Days[last7Days.length - 1];
    
    const startInput = getElement('startDate');
    const endInput = getElement('endDate');
    
    if (startInput) startInput.value = startDate;
    if (endInput) endInput.value = endDate;
    
    currentDateRange.startDate = startDate;
    currentDateRange.endDate = endDate;
    
    vehicleTabFilters.startDate = startDate;
    vehicleTabFilters.endDate = endDate;
    
    const mainStartInput = getElement('mainStartDate');
    const mainEndInput = getElement('mainEndDate');
    if (mainStartInput) mainStartInput.value = startDate;
    if (mainEndInput) mainEndInput.value = endDate;
    
    mainFilters.startDate = startDate;
    mainFilters.endDate = endDate;
    
    if (dataLoaded) {
        updateDashboard();
    }
}

// Apply custom date range
function applyDateRange() {
    const startInput = getElement('startDate');
    const endInput = getElement('endDate');
    
    if (!startInput || !endInput || !startInput.value || !endInput.value) {
        showNotification('Please select both start and end dates', 'warning');
        return;
    }
    
    if (startInput.value > endInput.value) {
        showNotification('Start date must be before end date', 'warning');
        return;
    }
    
    currentDateRange.startDate = startInput.value;
    currentDateRange.endDate = endInput.value;
    
    vehicleTabFilters.startDate = startInput.value;
    vehicleTabFilters.endDate = endInput.value;
    
    const mainStartInput = getElement('mainStartDate');
    const mainEndInput = getElement('mainEndDate');
    if (mainStartInput) mainStartInput.value = startInput.value;
    if (mainEndInput) mainEndInput.value = endInput.value;
    
    mainFilters.startDate = startInput.value;
    mainFilters.endDate = endInput.value;
    
    if (dataLoaded) {
        updateDashboard();
    }
}

// Get all dates in current range
function getDateRange() {
    if (!currentDateRange.startDate || !currentDateRange.endDate) {
        setLast7Days();
        return getLast7Days();
    }
    
    const dates = [];
    const start = new Date(currentDateRange.startDate);
    const end = new Date(currentDateRange.endDate);
    
    for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        dates.push(`${year}-${month}-${day}`);
    }
    
    return dates;
}

// Format date range for display
function formatDateRange(dates) {
    if (!dates || dates.length === 0) return 'No dates';
    const firstDate = new Date(dates[0]);
    const lastDate = new Date(dates[dates.length - 1]);
    
    const options = { month: 'short', day: 'numeric', year: 'numeric' };
    return `${firstDate.toLocaleDateString('en-US', options)} - ${lastDate.toLocaleDateString('en-US', options)}`;
}

// Improved CSV parser that handles quoted fields properly
function parseCSV(csvText) {
    const rows = [];
    const lines = csvText.split('\n');
    
    for (let line of lines) {
        if (line.trim() === '') continue;
        
        const row = [];
        let inQuote = false;
        let currentValue = '';
        
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            
            if (char === '"') {
                inQuote = !inQuote;
            } else if (char === ',' && !inQuote) {
                row.push(currentValue.trim());
                currentValue = '';
            } else {
                currentValue += char;
            }
        }
        
        // Add the last value
        row.push(currentValue.trim());
        
        // Clean up quoted values
        const cleanedRow = row.map(val => {
            val = val.trim();
            if (val.startsWith('"') && val.endsWith('"')) {
                val = val.substring(1, val.length - 1);
            }
            return val;
        });
        
        rows.push(cleanedRow);
    }
    
    return rows;
}

async function loadAllSheets() {
    const loadingEl = getElement('loading');
    const errorEl = getElement('error');
    
    if (errorEl) {
        errorEl.classList.remove('active');
        errorEl.textContent = '';
    }
    
    // Run date test on load
    testDateNormalization();
    
    try {
        allSheetData = {};
        let loadedCount = 0;
        
        for (const [tabName, tabInfo] of Object.entries(TABS)) {
            try {
                console.log(`Loading ${tabName}...`);
                const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${tabInfo.gid}`;
                const response = await fetch(url);
                
                if (!response.ok) {
                    throw new Error(`HTTP ${response.status}`);
                }
                
                const csvText = await response.text();
                
                if (!csvText || csvText.trim() === '') {
                    throw new Error('Empty response');
                }
                
                const rows = parseCSV(csvText);
                
                if (rows.length === 0) {
                    throw new Error('No data rows');
                }
                
                // Get headers from first row
                const rawHeaders = rows[0] || [];
                
                // Filter valid headers and get their indices
                const validIndices = [];
                const validHeaders = [];
                
                rawHeaders.forEach((header, idx) => {
                    if (isValidHeader(header) && !isColumnNumberHeader(header)) {
                        validIndices.push(idx);
                        validHeaders.push(header.trim());
                    }
                });
                
                // Process data rows (skip header row)
                const dataRows = rows.slice(1)
                    .filter(row => row && row.length > 0 && row.some(cell => cell && cell.trim() !== ''))
                    .map(row => validIndices.map(idx => (row[idx] || '').trim()));
                
                allSheetData[tabName] = {
                    headers: validHeaders,
                    rawHeaders: validHeaders,
                    data: dataRows,
                    type: tabInfo.type
                };
                
                loadedCount++;
                console.log(`✅ Loaded ${tabName}: ${validHeaders.length} columns, ${dataRows.length} rows`);
                
            } catch (error) {
                console.error(`❌ Error loading ${tabName}:`, error);
                allSheetData[tabName] = { 
                    headers: [], 
                    rawHeaders: [], 
                    data: [], 
                    type: tabInfo.type 
                };
            }
        }
        
        if (loadedCount === 0) {
            throw new Error('Failed to load any sheets');
        }
        
        // Process data
        processVehicleRecords();
        processDailyTotals();
        
        const footerTimestamp = getElement('footerTimestamp');
        if (footerTimestamp) {
            footerTimestamp.textContent = new Date().toLocaleString();
        }
        
        populateMainFilters();
        setLast7Days();
        
        dataLoaded = true;
        switchToView('dashboard');
        showNotification('Data loaded successfully!', 'success');
        
    } catch (error) {
        console.error('Fatal error:', error);
        if (errorEl) {
            errorEl.textContent = 'Error loading data: ' + error.message;
            errorEl.classList.add('active');
        }
        showNotification('Failed to load data', 'error');
    } finally {
        if (loadingEl) {
            loadingEl.classList.add('hidden');
        }
    }
}

function processVehicleRecords() {
    vehicleRecords = { 'BD78NGZN': [], 'CS44GHNZ': [], 'DG28ZLZN': [] };
    
    VEHICLES.forEach(vehicle => {
        const vehicleData = allSheetData[vehicle];
        if (!vehicleData || !vehicleData.data || vehicleData.data.length === 0) {
            console.warn(`No data for ${vehicle}`);
            return;
        }
        
        const headers = vehicleData.rawHeaders;
        
        // Find column indices - try different possible column names
        let dateIdx = findColumnIndex(headers, 'stop time');
        if (dateIdx === -1) dateIdx = findColumnIndex(headers, 'date');
        if (dateIdx === -1) dateIdx = findColumnIndex(headers, 'timestamp');
        
        let distIdx = findColumnIndex(headers, 'dist');
        if (distIdx === -1) distIdx = findColumnIndex(headers, 'distance');
        if (distIdx === -1) distIdx = findColumnIndex(headers, 'km');
        
        if (dateIdx === -1) {
            console.warn(`Date column not found in ${vehicle}. Headers:`, headers);
            return;
        }
        
        if (distIdx === -1) {
            console.warn(`Distance column not found in ${vehicle}. Headers:`, headers);
            return;
        }
        
        console.log(`Processing ${vehicle}: Date col ${dateIdx} (${headers[dateIdx]}), Dist col ${distIdx} (${headers[distIdx]})`);
        
        vehicleData.data.forEach(row => {
            const rawDate = row[dateIdx] || '';
            // Pass the vehicle name to normalizeDate for proper format detection
            const date = normalizeDate(rawDate, vehicle);
            
            if (!date) return;
            
            // Parse distance - handle various formats
            let distance = 0;
            const distStr = (row[distIdx] || '').replace(/[^\d.,-]/g, '').replace(',', '.');
            const parsed = parseFloat(distStr);
            if (!isNaN(parsed) && parsed > 0.10) {
                distance = parsed;
            }
            
            if (distance > 0) {
                vehicleRecords[vehicle].push({
                    date: date,
                    distance: distance,
                    vehicle: vehicle
                });
            }
        });
        
        // Sort by date
        vehicleRecords[vehicle].sort((a, b) => a.date.localeCompare(b.date));
        console.log(`${vehicle}: ${vehicleRecords[vehicle].length} records with distance > 0`);
        
        // Log sample dates to verify normalization
        if (vehicleRecords[vehicle].length > 0) {
            console.log(`${vehicle} sample dates:`, vehicleRecords[vehicle].slice(0, 3).map(r => r.date));
        }
    });
}

function processDailyTotals() {
    dailyTotals = { MAIN: {}, VEHICLE: {} };
    
    VEHICLES.forEach(vehicle => {
        if (!dailyTotals.MAIN[vehicle]) dailyTotals.MAIN[vehicle] = {};
        if (!dailyTotals.VEHICLE[vehicle]) dailyTotals.VEHICLE[vehicle] = {};
    });
    
    // Process MAIN data
    const mainData = allSheetData['MAIN'];
    if (mainData && mainData.data && mainData.data.length > 0) {
        const headers = mainData.rawHeaders;
        
        // Find column indices
        let regIdx = findColumnIndex(headers, 'truck');
        if (regIdx === -1) regIdx = findColumnIndex(headers, 'reg');
        if (regIdx === -1) regIdx = findColumnIndex(headers, 'vehicle');
        
        let dateIdx = findColumnIndex(headers, 'date');
        if (dateIdx === -1) dateIdx = findColumnIndex(headers, 'timestamp');
        if (dateIdx === -1) dateIdx = findColumnIndex(headers, 'day');
        
        let startIdx = findColumnIndex(headers, 'start');
        let endIdx = findColumnIndex(headers, 'end');
        
        if (regIdx !== -1 && dateIdx !== -1 && startIdx !== -1 && endIdx !== -1) {
            console.log(`Processing MAIN: Reg col ${regIdx} (${headers[regIdx]}), Date col ${dateIdx} (${headers[dateIdx]})`);
            
            mainData.data.forEach(row => {
                const reg = row[regIdx] || '';
                const vehicle = extractRegistration(reg);
                
                if (vehicle && VEHICLES.includes(vehicle)) {
                    const rawDate = row[dateIdx] || '';
                    // Pass 'MAIN' to normalizeDate for proper format detection
                    const date = normalizeDate(rawDate, 'MAIN');
                    
                    if (!date) return;
                    
                    // Parse odometer values - handle various formats
                    const start = parseFloat(String(row[startIdx] || '0').replace(/[^\d.-]/g, '')) || 0;
                    const end = parseFloat(String(row[endIdx] || '0').replace(/[^\d.-]/g, '')) || 0;
                    
                    // Calculate distance (ensure it's positive)
                    const distance = Math.max(0, end - start);
                    
                    // Only include if distance is positive (greater than 0.1 km to filter out zero/invalid entries)
                    if (distance > 0.1) {
                        if (!dailyTotals.MAIN[vehicle][date]) {
                            dailyTotals.MAIN[vehicle][date] = 0;
                        }
                        dailyTotals.MAIN[vehicle][date] += distance;
                        
                        // Log for debugging
                        console.log(`MAIN: ${vehicle} on ${date}: start=${start}, end=${end}, distance=${distance}`);
                    }
                }
            });
        } else {
            console.warn('Required columns not found in MAIN:', { regIdx, dateIdx, startIdx, endIdx });
        }
    }
    
    // Process Vehicle tabs
    VEHICLES.forEach(vehicle => {
        (vehicleRecords[vehicle] || []).forEach(record => {
            if (record.distance > 0) {
                const date = record.date;
                if (!dailyTotals.VEHICLE[vehicle][date]) {
                    dailyTotals.VEHICLE[vehicle][date] = 0;
                }
                dailyTotals.VEHICLE[vehicle][date] += record.distance;
            }
        });
    });
    
    // Log totals for debugging
    VEHICLES.forEach(vehicle => {
        const mainEntries = Object.keys(dailyTotals.MAIN[vehicle] || {}).length;
        const vehicleEntries = Object.keys(dailyTotals.VEHICLE[vehicle] || {}).length;
        console.log(`${vehicle}: MAIN days: ${mainEntries}, TAB days: ${vehicleEntries}`);
        
        // Log sample totals to verify values
        if (mainEntries > 0) {
            const mainSample = Object.entries(dailyTotals.MAIN[vehicle]).slice(0, 3);
            console.log(`${vehicle} MAIN sample:`, mainSample);
        }
        if (vehicleEntries > 0) {
            const vehicleSample = Object.entries(dailyTotals.VEHICLE[vehicle]).slice(0, 3);
            console.log(`${vehicle} TAB sample:`, vehicleSample);
        }
    });
}

function filterVehicleRows(rows, tabName) {
    const tabData = allSheetData[tabName];
    if (!tabData || !tabData.rawHeaders) return rows;
    
    const headers = tabData.rawHeaders;
    
    // Find date column
    let dateIdx = findColumnIndex(headers, 'stop time');
    if (dateIdx === -1) dateIdx = findColumnIndex(headers, 'date');
    if (dateIdx === -1) dateIdx = findColumnIndex(headers, 'timestamp');
    
    // Find distance column
    let distIdx = findColumnIndex(headers, 'dist');
    if (distIdx === -1) distIdx = findColumnIndex(headers, 'distance');
    if (distIdx === -1) distIdx = findColumnIndex(headers, 'km');
    
    return rows.filter(row => {
        // Apply date filter
        if ((vehicleTabFilters.startDate || vehicleTabFilters.endDate) && dateIdx !== -1) {
            const rowDate = normalizeDate(row[dateIdx] || '', tabName);
            if (rowDate) {
                if (vehicleTabFilters.startDate && rowDate < vehicleTabFilters.startDate) return false;
                if (vehicleTabFilters.endDate && rowDate > vehicleTabFilters.endDate) return false;
            }
        }
        
        // Filter out small distances
        if (distIdx !== -1 && row[distIdx]) {
            const distStr = row[distIdx].replace(/[^\d.,-]/g, '').replace(',', '.');
            const distance = parseFloat(distStr) || 0;
            if (distance <= 0.10) return false;
        }
        
        return true;
    });
}

function filterMainRows(rows) {
    const headers = allSheetData['MAIN']?.rawHeaders || [];
    
    let driverIdx = findColumnIndex(headers, 'driver');
    let truckIdx = findColumnIndex(headers, 'truck');
    if (truckIdx === -1) truckIdx = findColumnIndex(headers, 'reg');
    let dateIdx = findColumnIndex(headers, 'date');
    if (dateIdx === -1) dateIdx = findColumnIndex(headers, 'timestamp');
    
    return rows.filter(row => {
        // Driver filter
        if (mainFilters.driver !== 'all' && driverIdx !== -1) {
            if ((row[driverIdx] || '').trim() !== mainFilters.driver) return false;
        }
        
        // Truck filter
        if (mainFilters.truck !== 'all' && truckIdx !== -1) {
            if ((row[truckIdx] || '').trim() !== mainFilters.truck) return false;
        }
        
        // Date filter
        if ((mainFilters.startDate || mainFilters.endDate) && dateIdx !== -1) {
            const rowDate = normalizeDate(row[dateIdx] || '', 'MAIN');
            if (rowDate) {
                if (mainFilters.startDate && rowDate < mainFilters.startDate) return false;
                if (mainFilters.endDate && rowDate > mainFilters.endDate) return false;
            }
        }
        
        return true;
    });
}

function calculateRowTotal(row, headers) {
    let startIdx = findColumnIndex(headers, 'start');
    let endIdx = findColumnIndex(headers, 'end');
    
    if (startIdx !== -1 && endIdx !== -1 && startIdx < row.length && endIdx < row.length) {
        // Parse odometer values, handling various formats
        const start = parseFloat(String(row[startIdx] || '0').replace(/[^\d.-]/g, '')) || 0;
        const end = parseFloat(String(row[endIdx] || '0').replace(/[^\d.-]/g, '')) || 0;
        const distance = Math.max(0, end - start);
        
        // Only return if distance is positive and reasonable (filter out zero/invalid)
        if (distance > 0.1) {
            return distance;
        }
    }
    return null;
}

function calculateVehicleTabTotal(rows, headers) {
    let distIdx = findColumnIndex(headers, 'dist');
    if (distIdx === -1) distIdx = findColumnIndex(headers, 'distance');
    if (distIdx === -1) distIdx = findColumnIndex(headers, 'km');
    
    if (distIdx === -1) return 0;
    
    let total = 0;
    rows.forEach(row => {
        if (row[distIdx]) {
            const distStr = row[distIdx].replace(/[^\d.,-]/g, '').replace(',', '.');
            const distance = parseFloat(distStr) || 0;
            if (distance > 0.10) {
                total += distance;
            }
        }
    });
    
    return total;
}

function calculateGrandTotal(rows, headers) {
    let total = 0;
    rows.forEach(row => {
        const rowTotal = calculateRowTotal(row, headers);
        if (rowTotal !== null) {
            total += rowTotal;
        }
    });
    return total;
}

function populateMainFilters() {
    const mainData = allSheetData['MAIN'];
    if (!mainData || !mainData.data) return;
    
    const headers = mainData.rawHeaders;
    
    let driverIdx = findColumnIndex(headers, 'driver');
    let truckIdx = findColumnIndex(headers, 'truck');
    if (truckIdx === -1) truckIdx = findColumnIndex(headers, 'reg');
    
    if (driverIdx !== -1) {
        const drivers = new Set();
        mainData.data.forEach(row => {
            if (row[driverIdx] && row[driverIdx].trim() !== '') {
                drivers.add(row[driverIdx].trim());
            }
        });
        
        const driverSelect = getElement('driverFilter');
        if (driverSelect) {
            const sortedDrivers = Array.from(drivers).sort();
            driverSelect.innerHTML = '<option value="all">All Drivers</option>' + 
                sortedDrivers.map(d => `<option value="${d}">${d}</option>`).join('');
        }
    }
    
    if (truckIdx !== -1) {
        const trucks = new Set();
        mainData.data.forEach(row => {
            if (row[truckIdx] && row[truckIdx].trim() !== '') {
                trucks.add(row[truckIdx].trim());
            }
        });
        
        const truckSelect = getElement('truckFilter');
        if (truckSelect) {
            const sortedTrucks = Array.from(trucks).sort();
            truckSelect.innerHTML = '<option value="all">All Trucks</option>' + 
                sortedTrucks.map(t => `<option value="${t}">${t}</option>`).join('');
        }
    }
}

function applyMainFilters() {
    const driverSelect = getElement('driverFilter');
    const truckSelect = getElement('truckFilter');
    const startDateInput = getElement('mainStartDate');
    const endDateInput = getElement('mainEndDate');
    
    mainFilters.driver = driverSelect ? driverSelect.value : 'all';
    mainFilters.truck = truckSelect ? truckSelect.value : 'all';
    mainFilters.startDate = startDateInput ? startDateInput.value : '';
    mainFilters.endDate = endDateInput ? endDateInput.value : '';
    
    // Update vehicle tab filters to match
    vehicleTabFilters.startDate = mainFilters.startDate;
    vehicleTabFilters.endDate = mainFilters.endDate;
    
    // Update dashboard date inputs if we're in dashboard view
    if (currentView === 'dashboard') {
        const startInput = getElement('startDate');
        const endInput = getElement('endDate');
        if (startInput) startInput.value = mainFilters.startDate;
        if (endInput) endInput.value = mainFilters.endDate;
        currentDateRange.startDate = mainFilters.startDate;
        currentDateRange.endDate = mainFilters.endDate;
        updateDashboard();
    } else if (currentView === 'MAIN') {
        // If we're on MAIN tab, refresh the display
        displayTabData('MAIN');
    }
}

function clearMainFilters() {
    const driverSelect = getElement('driverFilter');
    const truckSelect = getElement('truckFilter');
    const startDateInput = getElement('mainStartDate');
    const endDateInput = getElement('mainEndDate');
    
    if (driverSelect) driverSelect.value = 'all';
    if (truckSelect) truckSelect.value = 'all';
    if (startDateInput) startDateInput.value = '';
    if (endDateInput) endDateInput.value = '';
    
    mainFilters = { driver: 'all', truck: 'all', startDate: '', endDate: '' };
    vehicleTabFilters = { startDate: '', endDate: '' };
    
    if (currentView === 'dashboard') {
        setLast7Days();
    } else if (currentView === 'MAIN') {
        displayTabData('MAIN');
    }
}

function displayTabData(tabName) {
    if (!dataLoaded) return;
    
    const tabData = allSheetData[tabName];
    if (!tabData) {
        console.warn(`No data for tab: ${tabName}`);
        return;
    }
    
    const excelGrid = getElement('tabExcelGrid');
    const tabTitle = getElement('currentTabTitle');
    const filterBar = getElement('mainFilterBar');
    const filterStats = getElement('mainFilterStats');
    const totalsContainer = getElement('totalsContainer');
    
    if (!excelGrid || !totalsContainer) return;
    
    excelGrid.innerHTML = '';
    if (tabTitle) tabTitle.textContent = tabName;
    
    if (filterBar) {
        filterBar.style.display = tabName === 'MAIN' ? 'flex' : 'none';
    }
    
    const headers = tabData.rawHeaders || [];
    
    // Find column indices for special handling
    let startIdx = findColumnIndex(headers, 'start');
    let endIdx = findColumnIndex(headers, 'end');
    let distIdx = findColumnIndex(headers, 'dist');
    if (distIdx === -1) distIdx = findColumnIndex(headers, 'distance');
    if (distIdx === -1) distIdx = findColumnIndex(headers, 'km');
    let fromIdx = findColumnIndex(headers, 'from');
    let toIdx = findColumnIndex(headers, 'to');
    
    // Create header row
    const headerRow = document.createElement('div');
    headerRow.className = 'excel-row-header';
    
    headers.forEach(header => {
        if (!header || header.trim() === '') return;
        
        const colHeader = document.createElement('div');
        colHeader.className = 'excel-col-header';
        
        // Choose icon based on header name
        let icon = 'fa-columns';
        const headerLower = header.toLowerCase();
        if (headerLower.includes('date')) icon = 'fa-calendar';
        else if (headerLower.includes('time')) icon = 'fa-clock';
        else if (headerLower.includes('driver')) icon = 'fa-user';
        else if (headerLower.includes('truck') || headerLower.includes('reg') || headerLower.includes('vehicle')) icon = 'fa-truck';
        else if (headerLower.includes('from')) icon = 'fa-map-marker-alt';
        else if (headerLower.includes('to')) icon = 'fa-map-pin';
        else if (headerLower.includes('odo') || headerLower.includes('start') || headerLower.includes('end')) icon = 'fa-tachometer-alt';
        else if (headerLower.includes('dist') || headerLower.includes('km')) icon = 'fa-road';
        else if (headerLower.includes('coord')) icon = 'fa-globe';
        else if (headerLower.includes('email')) icon = 'fa-envelope';
        else if (headerLower.includes('notes')) icon = 'fa-sticky-note';
        else if (headerLower.includes('file')) icon = 'fa-file';
        else if (headerLower.includes('contact')) icon = 'fa-address-book';
        
        colHeader.innerHTML = `<i class="fas ${icon}"></i> ${header}`;
        headerRow.appendChild(colHeader);
    });
    
    // Add total column for MAIN tab
    if (tabName === 'MAIN' && startIdx !== -1 && endIdx !== -1) {
        const totalHeader = document.createElement('div');
        totalHeader.className = 'excel-col-header total-column';
        totalHeader.innerHTML = `<i class="fas fa-calculator"></i> Distance (km)`;
        headerRow.appendChild(totalHeader);
    }
    
    excelGrid.appendChild(headerRow);
    
    // Get filtered data
    let rowsToDisplay;
    let hiddenCount = 0;
    
    if (tabName === 'MAIN') {
        rowsToDisplay = filterMainRows(tabData.data);
        
        if (filterStats) {
            filterStats.textContent = `Showing ${rowsToDisplay.length} of ${tabData.data.length} rows`;
        }
        
        // Calculate grand total
        let grandTotal = 0;
        let validRowCount = 0;
        
        rowsToDisplay.forEach(row => {
            const rowTotal = calculateRowTotal(row, headers);
            if (rowTotal !== null) {
                grandTotal += rowTotal;
                validRowCount++;
            }
        });
        
        // Count rows with zero/invalid distances
        const zeroDistanceRows = rowsToDisplay.filter(row => {
            const rowTotal = calculateRowTotal(row, headers);
            return rowTotal === null || rowTotal <= 0.1;
        }).length;
        
        totalsContainer.innerHTML = `
            <div style="display: flex; flex-direction: column; gap: 5px;">
                <div>
                    <span class="totals-label">
                        <i class="fas fa-calculator"></i> Total Distance:
                    </span>
                    <span class="totals-value">${Math.round(grandTotal).toLocaleString()} km</span>
                </div>
                <div style="font-size: 12px; color: #666;">
                    <span><i class="fas fa-check-circle" style="color: #2e7d32;"></i> Valid rows: ${validRowCount}</span>
                    <span style="margin-left: 15px;"><i class="fas fa-info-circle" style="color: #ed6c02;"></i> Zero distance: ${zeroDistanceRows}</span>
                </div>
            </div>
        `;
        
    } else {
        // For vehicle tabs
        const allRows = tabData.data;
        rowsToDisplay = filterVehicleRows(allRows, tabName);
        
        if (distIdx !== -1) {
            hiddenCount = allRows.filter(row => {
                if (row[distIdx]) {
                    const distStr = row[distIdx].replace(/[^\d.,-]/g, '').replace(',', '.');
                    const distance = parseFloat(distStr) || 0;
                    return distance <= 0.10 && distance > 0;
                }
                return false;
            }).length;
        }
        
        if (filterStats) {
            filterStats.innerHTML = `Showing ${rowsToDisplay.length} of ${tabData.data.length} rows`;
            if (distIdx !== -1 && hiddenCount > 0) {
                filterStats.innerHTML += `<span style="margin-left:15px; color:#1976d2;"><i class="fas fa-filter"></i> Hidden (Dist ≤ 0.10): ${hiddenCount} rows</span>`;
            }
        }
        
        const vehicleTotal = calculateVehicleTabTotal(rowsToDisplay, headers);
        
        totalsContainer.innerHTML = `
            <span class="totals-label">
                <i class="fas fa-road"></i> Total Distance:
            </span>
            <span class="totals-value">${Math.round(vehicleTotal).toLocaleString()} km</span>
            <div class="totals-breakdown">
                <span class="totals-breakdown-item">
                    <i class="fas fa-eye"></i> Visible: ${rowsToDisplay.length}
                </span>
                <span class="totals-breakdown-item">
                    <i class="fas fa-filter"></i> Hidden: ${hiddenCount}
                </span>
            </div>
        `;
    }
    
    // Create data rows
    rowsToDisplay.forEach(row => {
        const excelRow = document.createElement('div');
        excelRow.className = 'excel-row';
        
        headers.forEach((header, colIndex) => {
            if (!header || header.trim() === '') return;
            
            const cell = document.createElement('div');
            cell.className = 'excel-cell';
            
            // Get cell value
            let cellValue = colIndex < row.length ? row[colIndex] : '-';
            
            // Highlight odometer readings in MAIN tab
            if (tabName === 'MAIN') {
                if (colIndex === startIdx || colIndex === endIdx) {
                    cell.classList.add('odo-highlight');
                    // Format odometer values nicely
                    if (cellValue && cellValue !== '-') {
                        const num = parseFloat(String(cellValue).replace(/[^\d.-]/g, ''));
                        if (!isNaN(num)) {
                            cellValue = Math.round(num).toLocaleString();
                        }
                    }
                }
                // Highlight location fields
                if (colIndex === fromIdx || colIndex === toIdx) {
                    cell.classList.add('location-highlight');
                }
            }
            
            // Highlight distance in vehicle tabs
            if (tabName !== 'MAIN' && colIndex === distIdx) {
                const distStr = row[colIndex] ? row[colIndex].replace(/[^\d.,-]/g, '').replace(',', '.') : '0';
                const distance = parseFloat(distStr) || 0;
                if (distance > 0.10) {
                    cell.style.fontWeight = '500';
                    cell.style.color = '#2e7d32';
                }
            }
            
            cell.textContent = cellValue;
            excelRow.appendChild(cell);
        });
        
        // Add total cell for MAIN tab with calculated distance
        if (tabName === 'MAIN' && startIdx !== -1 && endIdx !== -1) {
            const totalCell = document.createElement('div');
            totalCell.className = 'excel-cell total-cell';
            
            const rowTotal = calculateRowTotal(row, headers);
            if (rowTotal !== null) {
                totalCell.textContent = Math.round(rowTotal).toLocaleString() + ' km';
                totalCell.style.fontWeight = '600';
                totalCell.style.color = '#1976d2';
            } else {
                totalCell.textContent = '-';
            }
            excelRow.appendChild(totalCell);
        }
        
        excelGrid.appendChild(excelRow);
    });
    
    // Show empty message if no data
    if (rowsToDisplay.length === 0) {
        const emptyRow = document.createElement('div');
        emptyRow.className = 'excel-row';
        emptyRow.style.padding = '20px';
        emptyRow.style.textAlign = 'center';
        emptyRow.style.color = '#666';
        emptyRow.style.justifyContent = 'center';
        
        if (tabName !== 'MAIN' && hiddenCount > 0) {
            emptyRow.innerHTML = `<i class="fas fa-info-circle"></i> No rows with distance > 0.10 km`;
        } else {
            emptyRow.innerHTML = `<i class="fas fa-info-circle"></i> No data matches filters`;
        }
        
        excelGrid.appendChild(emptyRow);
    }
    
    // Add grand total row for MAIN tab
    if (tabName === 'MAIN' && rowsToDisplay.length > 0) {
        const grandTotalRow = document.createElement('div');
        grandTotalRow.className = 'excel-row grand-total-row';
        
        // Add empty cells for each header
        headers.forEach(header => {
            if (!header || header.trim() === '') return;
            const cell = document.createElement('div');
            cell.className = 'excel-cell';
            cell.textContent = '';
            grandTotalRow.appendChild(cell);
        });
        
        // Add grand total cell
        const grandTotalCell = document.createElement('div');
        grandTotalCell.className = 'excel-cell grand-total-cell';
        const grandTotal = calculateGrandTotal(rowsToDisplay, headers);
        grandTotalCell.innerHTML = `<strong>${Math.round(grandTotal).toLocaleString()} km</strong>`;
        grandTotalRow.appendChild(grandTotalCell);
        
        excelGrid.appendChild(grandTotalRow);
    }
    
    const footerVehicleCount = getElement('footerVehicleCount');
    if (footerVehicleCount) footerVehicleCount.textContent = rowsToDisplay.length;
}

function updateDashboard() {
    if (!dataLoaded) return;
    
    const selectedVehicle = getElement('vehicleFilter')?.value || 'all';
    const dateRange = getDateRange();
    
    const dateRangeText = getElement('dateRangeText');
    const footerDateRange = getElement('footerDateRange');
    const formattedRange = formatDateRange(dateRange);
    if (dateRangeText) dateRangeText.textContent = formattedRange;
    if (footerDateRange) footerDateRange.textContent = formattedRange;
    
    const vehiclesToShow = selectedVehicle === 'all' ? VEHICLES : [selectedVehicle];
    
    // Update chart
    const chartData = prepareDateRangeData(vehiclesToShow, dateRange);
    updateChart(chartData);
    
    // Update stats cards with REAL calculated values
    const statsContainer = getElement('vehicleStatsCards');
    if (statsContainer) {
        statsContainer.innerHTML = '';
        
        vehiclesToShow.forEach(vehicle => {
            let mainTotal = 0;
            let vehicleTotal = 0;
            
            dateRange.forEach(date => {
                if (dailyTotals.MAIN[vehicle] && dailyTotals.MAIN[vehicle][date]) {
                    mainTotal += dailyTotals.MAIN[vehicle][date];
                }
                if (dailyTotals.VEHICLE[vehicle] && dailyTotals.VEHICLE[vehicle][date]) {
                    vehicleTotal += dailyTotals.VEHICLE[vehicle][date];
                }
            });
            
            const variance = mainTotal - vehicleTotal;
            const variancePercent = mainTotal > 0 ? ((variance / mainTotal) * 100).toFixed(1) : '0';
            
            const card = document.createElement('div');
            card.className = 'vehicle-card';
            
            card.innerHTML = `
                <div class="vehicle-title">
                    <i class="fas fa-truck"></i> ${vehicle}
                </div>
                <div class="vehicle-stats">
                    <div class="stat-group">
                        <div class="stat-label">MAIN (Odometer)</div>
                        <div class="stat-number main">${Math.round(mainTotal).toLocaleString()}</div>
                    </div>
                    <div class="stat-group">
                        <div class="stat-label">TRACKER (GPS)</div>
                        <div class="stat-number tracker">${Math.round(vehicleTotal).toLocaleString()}</div>
                    </div>
                </div>
                <span class="variance-badge ${variance >= 0 ? 'positive' : 'negative'}">
                    ${variance >= 0 ? '+' : ''}${Math.round(variance).toLocaleString()} (${variancePercent}%)
                </span>
            `;
            
            statsContainer.appendChild(card);
        });
    }
    
    // Update comparison table
    updateComparisonTable(vehiclesToShow, dateRange);
    
    const footerVehicleCount = getElement('footerVehicleCount');
    if (footerVehicleCount) footerVehicleCount.textContent = vehiclesToShow.length;
}

function prepareDateRangeData(vehicles, dateRange) {
    const colors = ['#4da6ff', '#ff9f4a', '#6c5ce7', '#00b894', '#e17055', '#0984e3'];
    
    return {
        labels: dateRange.map(date => {
            const d = new Date(date);
            return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
        }),
        datasets: vehicles.flatMap((vehicle, idx) => {
            const mainData = dateRange.map(date => Math.round(dailyTotals.MAIN[vehicle]?.[date] || 0));
            const vehicleData = dateRange.map(date => Math.round(dailyTotals.VEHICLE[vehicle]?.[date] || 0));
            const color = colors[idx % colors.length];
            
            return [
                {
                    label: `${vehicle} (MAIN)`,
                    data: mainData,
                    borderColor: color,
                    backgroundColor: 'transparent',
                    borderWidth: 1.2,
                    borderDash: [],
                    tension: 0.2,
                    fill: false,
                    pointRadius: 2,
                    pointHoverRadius: 3
                },
                {
                    label: `${vehicle} (TRACKER)`,
                    data: vehicleData,
                    borderColor: color,
                    backgroundColor: 'transparent',
                    borderWidth: 1.2,
                    borderDash: [4, 3],
                    tension: 0.2,
                    fill: false,
                    pointRadius: 2,
                    pointHoverRadius: 3
                }
            ];
        })
    };
}

function updateChart(chartData) {
    const canvas = getElement('distanceChart');
    if (!canvas) return;
    
    if (distanceChart) {
        distanceChart.destroy();
    }
    
    const legend = getElement('chartLegend');
    if (legend) {
        legend.innerHTML = '';
        chartData.datasets.forEach(dataset => {
            legend.innerHTML += `
                <div class="legend-item">
                    <span class="legend-color" style="background-color: ${dataset.borderColor}"></span>
                    <span>${dataset.label}</span>
                </div>
            `;
        });
    }
    
    try {
        distanceChart = new Chart(canvas, {
            type: 'line',
            data: chartData,
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                plugins: {
                    tooltip: {
                        callbacks: {
                            label: (ctx) => `${ctx.dataset.label}: ${ctx.raw.toLocaleString()} km`
                        }
                    },
                    legend: { display: false }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        title: { display: true, text: 'km' },
                        ticks: { callback: v => v.toLocaleString() }
                    }
                }
            }
        });
    } catch (e) {
        console.error('Chart error:', e);
    }
}

function updateComparisonTable(vehicles, dateRange) {
    const headersContainer = getElement('comparisonHeaders');
    const rowsContainer = getElement('comparisonRows');
    
    if (!headersContainer || !rowsContainer) return;
    
    headersContainer.innerHTML = '';
    rowsContainer.innerHTML = '';
    
    const headers = ['Date'];
    vehicles.forEach(vehicle => {
        headers.push(`${vehicle} MAIN`, `${vehicle} TRACKER`, 'Var');
    });
    
    headers.forEach(header => {
        const el = document.createElement('div');
        el.className = 'comp-header';
        el.textContent = header;
        headersContainer.appendChild(el);
    });
    
    dateRange.forEach((date, index) => {
        const row = document.createElement('div');
        row.className = 'comp-row';
        
        let hasData = false;
        vehicles.forEach(vehicle => {
            if ((dailyTotals.MAIN[vehicle]?.[date] || 0) > 0 || 
                (dailyTotals.VEHICLE[vehicle]?.[date] || 0) > 0) {
                hasData = true;
            }
        });
        
        if (!hasData) {
            row.classList.add('zero-day');
        }
        
        const rowNum = document.createElement('div');
        rowNum.className = 'row-number';
        rowNum.textContent = index + 1;
        row.appendChild(rowNum);
        
        const dateCell = document.createElement('div');
        dateCell.className = 'comp-cell';
        dateCell.textContent = new Date(date).toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
        if (!hasData) dateCell.classList.add('zero-value');
        row.appendChild(dateCell);
        
        vehicles.forEach(vehicle => {
            const mainVal = dailyTotals.MAIN[vehicle]?.[date] || 0;
            const vehicleVal = dailyTotals.VEHICLE[vehicle]?.[date] || 0;
            const variance = mainVal - vehicleVal;
            
            const mainCell = document.createElement('div');
            mainCell.className = 'comp-cell';
            mainCell.textContent = mainVal > 0 ? Math.round(mainVal).toLocaleString() : '-';
            if (mainVal === 0) mainCell.classList.add('zero-value');
            row.appendChild(mainCell);
            
            const vehicleCell = document.createElement('div');
            vehicleCell.className = 'comp-cell';
            vehicleCell.textContent = vehicleVal > 0 ? Math.round(vehicleVal).toLocaleString() : '-';
            if (vehicleVal === 0) vehicleCell.classList.add('zero-value');
            row.appendChild(vehicleCell);
            
            const varCell = document.createElement('div');
            varCell.className = 'comp-cell';
            if (Math.abs(variance) > 10 && mainVal > 0 && vehicleVal > 0) {
                varCell.classList.add(variance > 0 ? 'positive' : 'negative');
            }
            varCell.textContent = variance !== 0 ? Math.round(variance).toLocaleString() : '-';
            if (variance === 0) varCell.classList.add('zero-value');
            row.appendChild(varCell);
        });
        
        rowsContainer.appendChild(row);
    });
}

function switchToView(view) {
    currentView = view;
    
    const dashboardView = getElement('dashboardView');
    const tableView = getElement('tableView');
    const vehicleSelector = getElement('vehicleSelector');
    const dateRangeSelector = getElement('dateRangeSelector');
    const dateRangeIndicator = getElement('dateRangeIndicator');
    const footerView = getElement('footerView');
    
    if (view === 'dashboard') {
        if (dashboardView) dashboardView.style.display = 'flex';
        if (tableView) tableView.style.display = 'none';
        if (vehicleSelector) vehicleSelector.style.display = 'flex';
        if (dateRangeSelector) dateRangeSelector.style.display = 'flex';
        if (dateRangeIndicator) dateRangeIndicator.style.display = 'flex';
        if (footerView) footerView.textContent = 'Dashboard';
        if (dataLoaded) updateDashboard();
    } else {
        if (dashboardView) dashboardView.style.display = 'none';
        if (tableView) {
            tableView.style.display = 'flex';
            if (dataLoaded) displayTabData(view);
        }
        if (vehicleSelector) vehicleSelector.style.display = 'none';
        if (dateRangeSelector) dateRangeSelector.style.display = 'none';
        if (dateRangeIndicator) dateRangeIndicator.style.display = 'none';
        if (footerView) footerView.textContent = view;
    }
}

function exportCurrentView() {
    if (currentView === 'dashboard') {
        exportComparisonData();
    } else {
        exportTabData();
    }
}

function exportTabData() {
    const tabData = allSheetData[currentView];
    if (!tabData || tabData.data.length === 0) {
        showNotification('No data to export', 'warning');
        return;
    }
    
    let dataToExport;
    if (currentView === 'MAIN') {
        dataToExport = filterMainRows(tabData.data);
    } else {
        dataToExport = filterVehicleRows(tabData.data, currentView);
    }
    
    const headers = [...(tabData.rawHeaders || [])];
    
    let html = '<html><head><meta charset="UTF-8"><title>' + currentView + '</title></head><body>';
    html += '<table border="1">';
    
    html += '<tr>';
    headers.forEach(header => {
        if (header && header.trim() !== '') {
            html += '<th>' + header + '</th>';
        }
    });
    
    if (currentView === 'MAIN') {
        html += '<th>Distance (km)</th>';
    } else {
        html += '<th>Distance (km)</th>';
    }
    html += '</tr>';
    
    dataToExport.forEach(row => {
        html += '<tr>';
        row.forEach(cell => {
            html += '<td>' + cell + '</td>';
        });
        
        if (currentView === 'MAIN') {
            let startIdx = findColumnIndex(headers, 'start');
            let endIdx = findColumnIndex(headers, 'end');
            
            if (startIdx !== -1 && endIdx !== -1 && startIdx < row.length && endIdx < row.length) {
                const start = parseFloat(String(row[startIdx] || '0').replace(/[^\d.-]/g, '')) || 0;
                const end = parseFloat(String(row[endIdx] || '0').replace(/[^\d.-]/g, '')) || 0;
                const total = Math.max(0, end - start);
                html += '<td>' + (total > 0.1 ? Math.round(total) : '-') + '</td>';
            } else {
                html += '<td>-</td>';
            }
        } else {
            let distIdx = findColumnIndex(headers, 'dist');
            if (distIdx === -1) distIdx = findColumnIndex(headers, 'distance');
            if (distIdx === -1) distIdx = findColumnIndex(headers, 'km');
            
            if (distIdx !== -1 && distIdx < row.length && row[distIdx]) {
                const distStr = row[distIdx].replace(/[^\d.,-]/g, '').replace(',', '.');
                const distance = parseFloat(distStr) || 0;
                html += '<td>' + (distance > 0.1 ? distance.toFixed(2) : '-') + '</td>';
            } else {
                html += '<td>-</td>';
            }
        }
        
        html += '</tr>';
    });
    
    html += '<tr style="font-weight:bold; background-color:#f0f0f0;">';
    html += '<td colspan="' + (headers.length) + '" style="text-align:right;">GRAND TOTAL:</td>';
    
    if (currentView === 'MAIN') {
        const grandTotal = calculateGrandTotal(dataToExport, headers);
        html += '<td>' + Math.round(grandTotal).toLocaleString() + ' km</td>';
    } else {
        const vehicleTotal = calculateVehicleTabTotal(dataToExport, headers);
        html += '<td>' + Math.round(vehicleTotal).toLocaleString() + ' km</td>';
    }
    
    html += '</tr>';
    html += '</table></body></html>';
    
    const blob = new Blob([html], { type: 'application/vnd.ms-excel' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.download = `${currentView}-${new Date().toISOString().slice(0,10)}.xls`;
    a.href = url;
    a.click();
    URL.revokeObjectURL(url);
    
    showNotification('Export successful!', 'success');
}

function exportComparisonData() {
    const selectedVehicle = getElement('vehicleFilter')?.value || 'all';
    const vehicles = selectedVehicle === 'all' ? VEHICLES : [selectedVehicle];
    const dateRange = getDateRange();
    
    let html = '<html><head><meta charset="UTF-8"><title>Daily Comparison</title></head><body>';
    html += '<table border="1">';
    
    html += '<tr><th>Date</th>';
    vehicles.forEach(vehicle => {
        html += `<th>${vehicle} MAIN (km)</th><th>${vehicle} TRACKER (km)</th><th>${vehicle} Variance</th>`;
    });
    html += '</tr>';
    
    dateRange.forEach(date => {
        html += '<tr>';
        html += `<td>${new Date(date).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}</td>`;
        
        vehicles.forEach(vehicle => {
            const mainVal = dailyTotals.MAIN[vehicle]?.[date] || 0;
            const vehicleVal = dailyTotals.VEHICLE[vehicle]?.[date] || 0;
            const variance = mainVal - vehicleVal;
            
            html += `<td>${Math.round(mainVal)}</td>`;
            html += `<td>${Math.round(vehicleVal)}</td>`;
            html += `<td>${Math.round(variance)}</td>`;
        });
        
        html += '</tr>';
    });
    
    html += '<tr style="font-weight:bold; background-color:#f0f0f0;">';
    html += '<td>GRAND TOTAL</td>';
    vehicles.forEach(vehicle => {
        let mainTotal = 0;
        let vehicleTotal = 0;
        
        dateRange.forEach(date => {
            mainTotal += dailyTotals.MAIN[vehicle]?.[date] || 0;
            vehicleTotal += dailyTotals.VEHICLE[vehicle]?.[date] || 0;
        });
        
        const variance = mainTotal - vehicleTotal;
        html += `<td>${Math.round(mainTotal)}</td>`;
        html += `<td>${Math.round(vehicleTotal)}</td>`;
        html += `<td>${Math.round(variance)}</td>`;
    });
    html += '</tr>';
    html += '</table></body></html>';
    
    const blob = new Blob([html], { type: 'application/vnd.ms-excel' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.download = `comparison-${selectedVehicle}-${new Date().toISOString().slice(0,10)}.xls`;
    a.href = url;
    a.click();
    URL.revokeObjectURL(url);
    
    showNotification('Export successful!', 'success');
}

function showNotification(message, type) {
    const notification = document.createElement('div');
    notification.className = 'excel-notification';
    notification.innerHTML = `<i class="fas ${type === 'success' ? 'fa-check-circle' : type === 'error' ? 'fa-exclamation-circle' : 'fa-info-circle'}"></i> ${message}`;
    document.body.appendChild(notification);
    
    setTimeout(() => {
        notification.style.opacity = '0';
        setTimeout(() => notification.remove(), 300);
    }, 3000);
}

// Initialize
document.addEventListener('DOMContentLoaded', function() {
    console.log('DOM ready, starting load...');
    loadAllSheets();
});