  // Configuration
    const SHEET_ID = '1UvzxLfth-MYeEa63k6VS1qpZfUXwKF_PholJ9YaDLO8';
    
    // Tab GIDs
    const TABS = {
        'MAIN': { gid: '1060733973', type: 'main' },
        'BD78NGZN': { gid: '1482391741', type: 'vehicle' },
        'CS44GHNZ': { gid: '416024164', type: 'vehicle' },
        'DG28ZLZN': { gid: '1908317780', type: 'vehicle' }
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

    // Helper: convert any date string to YYYY-MM-DD
    function normalizeDate(dateStr) {
        if (!dateStr) return '';
        try {
            let datePart = dateStr.split(' ')[0];
            let parts = datePart.split('/');
            if (parts.length === 3) {
                let month, day, year;
                // Try to determine if first part is month or day
                if (parseInt(parts[0]) > 12) {
                    // First part is likely day
                    day = parts[0].padStart(2, '0');
                    month = parts[1].padStart(2, '0');
                    year = parts[2];
                } else {
                    // First part is likely month
                    month = parts[0].padStart(2, '0');
                    day = parts[1].padStart(2, '0');
                    year = parts[2];
                }
                if (year.length === 2) year = '20' + year;
                return `${year}-${month}-${day}`;
            }
            return datePart;
        } catch (e) {
            console.warn('Date normalization error:', e);
            return '';
        }
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
        return headers.findIndex(header => 
            header && typeof header === 'string' && header.toLowerCase().includes(searchTerm.toLowerCase())
        );
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

    // Simple CSV parser
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
                    row.push(currentValue);
                    currentValue = '';
                } else {
                    currentValue += char;
                }
            }
            
            row.push(currentValue);
            rows.push(row);
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
        
        try {
            allSheetData = {};
            let loadedCount = 0;
            const totalTabs = Object.keys(TABS).length;
            
            for (const [tabName, tabInfo] of Object.entries(TABS)) {
                try {
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
                    
                    const rawHeaders = rows[0] || [];
                    
                    // Filter valid headers
                    const validIndices = [];
                    const validHeaders = [];
                    
                    rawHeaders.forEach((header, idx) => {
                        if (isValidHeader(header) && !isColumnNumberHeader(header)) {
                            validIndices.push(idx);
                            validHeaders.push(header.trim());
                        }
                    });
                    
                    // Process data rows
                    const dataRows = rows.slice(1)
                        .filter(row => row && row.some(cell => cell && cell.trim() !== ''))
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
            if (!vehicleData || !vehicleData.data) return;
            
            const headers = vehicleData.rawHeaders;
            const dateIdx = findColumnIndex(headers, 'stop time');
            const distIdx = findColumnIndex(headers, 'dist');
            
            if (dateIdx === -1 || distIdx === -1) {
                console.warn(`Required columns not found in ${vehicle}`);
                return;
            }
            
            vehicleData.data.forEach(row => {
                const date = normalizeDate(row[dateIdx] || '');
                if (!date) return;
                
                let distance = 0;
                const distStr = (row[distIdx] || '').replace(',', '.').replace(/[^0-9.-]/g, '');
                distance = parseFloat(distStr) || 0;
                
                vehicleRecords[vehicle].push({
                    date: date,
                    distance: distance,
                    vehicle: vehicle
                });
            });
            
            vehicleRecords[vehicle].sort((a, b) => a.date.localeCompare(b.date));
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
        if (mainData && mainData.data) {
            const headers = mainData.rawHeaders;
            const regIdx = findColumnIndex(headers, 'truck') !== -1 ? 
                findColumnIndex(headers, 'truck') : findColumnIndex(headers, 'reg');
            const dateIdx = findColumnIndex(headers, 'date') !== -1 ? 
                findColumnIndex(headers, 'date') : findColumnIndex(headers, 'timestamp');
            const startIdx = findColumnIndex(headers, 'start');
            const endIdx = findColumnIndex(headers, 'end');
            
            if (regIdx !== -1 && dateIdx !== -1 && startIdx !== -1 && endIdx !== -1) {
                mainData.data.forEach(row => {
                    const reg = row[regIdx] || '';
                    const vehicle = extractRegistration(reg);
                    
                    if (vehicle && VEHICLES.includes(vehicle)) {
                        const date = normalizeDate(row[dateIdx] || '');
                        if (!date) return;
                        
                        const start = parseFloat(row[startIdx]) || 0;
                        const end = parseFloat(row[endIdx]) || 0;
                        const distance = Math.max(0, end - start);
                        
                        if (!dailyTotals.MAIN[vehicle][date]) {
                            dailyTotals.MAIN[vehicle][date] = 0;
                        }
                        dailyTotals.MAIN[vehicle][date] += distance;
                    }
                });
            }
        }
        
        // Process Vehicle tabs
        VEHICLES.forEach(vehicle => {
            (vehicleRecords[vehicle] || []).forEach(record => {
                if (record.distance > 0.10) {
                    const date = record.date;
                    if (!dailyTotals.VEHICLE[vehicle][date]) {
                        dailyTotals.VEHICLE[vehicle][date] = 0;
                    }
                    dailyTotals.VEHICLE[vehicle][date] += record.distance;
                }
            });
        });
    }

    function filterVehicleRowsByDistance(rows, headers) {
        const distIdx = findColumnIndex(headers, 'dist');
        if (distIdx === -1) return rows;
        
        return rows.filter(row => {
            if (row[distIdx]) {
                const distStr = row[distIdx].replace(',', '.').replace(/[^0-9.-]/g, '');
                const distance = parseFloat(distStr) || 0;
                return distance > 0.10;
            }
            return true;
        });
    }

    function filterVehicleRows(rows) {
        const headers = allSheetData[currentView]?.rawHeaders || [];
        const dateIdx = findColumnIndex(headers, 'stop time');
        
        let filteredRows = rows;
        
        if ((vehicleTabFilters.startDate || vehicleTabFilters.endDate) && dateIdx !== -1) {
            filteredRows = filteredRows.filter(row => {
                const rowDate = normalizeDate(row[dateIdx] || '');
                if (!rowDate) return true;
                
                if (vehicleTabFilters.startDate && rowDate < vehicleTabFilters.startDate) return false;
                if (vehicleTabFilters.endDate && rowDate > vehicleTabFilters.endDate) return false;
                
                return true;
            });
        }
        
        filteredRows = filterVehicleRowsByDistance(filteredRows, headers);
        return filteredRows;
    }

    function filterMainRows(rows) {
        const headers = allSheetData['MAIN']?.rawHeaders || [];
        const driverIdx = findColumnIndex(headers, 'driver');
        const truckIdx = findColumnIndex(headers, 'truck') !== -1 ? 
            findColumnIndex(headers, 'truck') : findColumnIndex(headers, 'reg');
        const dateIdx = findColumnIndex(headers, 'date') !== -1 ? 
            findColumnIndex(headers, 'date') : findColumnIndex(headers, 'timestamp');
        
        return rows.filter(row => {
            if (mainFilters.driver !== 'all' && driverIdx !== -1) {
                if (row[driverIdx] !== mainFilters.driver) return false;
            }
            
            if (mainFilters.truck !== 'all' && truckIdx !== -1) {
                if (row[truckIdx] !== mainFilters.truck) return false;
            }
            
            if ((mainFilters.startDate || mainFilters.endDate) && dateIdx !== -1) {
                const rowDate = normalizeDate(row[dateIdx] || '');
                if (rowDate) {
                    if (mainFilters.startDate && rowDate < mainFilters.startDate) return false;
                    if (mainFilters.endDate && rowDate > mainFilters.endDate) return false;
                }
            }
            
            return true;
        });
    }

    function calculateRowTotal(row, headers) {
        const startIdx = findColumnIndex(headers, 'start');
        const endIdx = findColumnIndex(headers, 'end');
        
        if (startIdx !== -1 && endIdx !== -1 && row[startIdx] && row[endIdx]) {
            const start = parseFloat(row[startIdx]) || 0;
            const end = parseFloat(row[endIdx]) || 0;
            return Math.max(0, end - start);
        }
        return null;
    }

    function calculateVehicleTabTotal(rows, headers) {
        const distIdx = findColumnIndex(headers, 'dist');
        if (distIdx === -1) return 0;
        
        let total = 0;
        rows.forEach(row => {
            if (row[distIdx]) {
                const distStr = row[distIdx].replace(',', '.').replace(/[^0-9.-]/g, '');
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
        const driverIdx = findColumnIndex(headers, 'driver');
        const truckIdx = findColumnIndex(headers, 'truck') !== -1 ? 
            findColumnIndex(headers, 'truck') : findColumnIndex(headers, 'reg');
        
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
        
        vehicleTabFilters.startDate = mainFilters.startDate;
        vehicleTabFilters.endDate = mainFilters.endDate;
        
        if (currentView === 'dashboard') {
            const startInput = getElement('startDate');
            const endInput = getElement('endDate');
            if (startInput) startInput.value = mainFilters.startDate;
            if (endInput) endInput.value = mainFilters.endDate;
            currentDateRange.startDate = mainFilters.startDate;
            currentDateRange.endDate = mainFilters.endDate;
            updateDashboard();
        } else {
            displayTabData(currentView);
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
        } else {
            displayTabData(currentView);
        }
    }

    function displayTabData(tabName) {
        if (!dataLoaded) return;
        
        const tabData = allSheetData[tabName];
        if (!tabData) return;
        
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
        const startIdx = findColumnIndex(headers, 'start');
        const endIdx = findColumnIndex(headers, 'end');
        const distIdx = findColumnIndex(headers, 'dist');
        
        // Create header row
        const headerRow = document.createElement('div');
        headerRow.className = 'excel-row-header';
        
        headers.forEach(header => {
            if (!header || header.trim() === '') return;
            
            const colHeader = document.createElement('div');
            colHeader.className = 'excel-col-header';
            
            let icon = 'fa-columns';
            const headerLower = header.toLowerCase();
            if (headerLower.includes('date')) icon = 'fa-calendar';
            else if (headerLower.includes('time')) icon = 'fa-clock';
            else if (headerLower.includes('driver')) icon = 'fa-user';
            else if (headerLower.includes('truck') || headerLower.includes('reg')) icon = 'fa-truck';
            else if (headerLower.includes('from')) icon = 'fa-map-marker-alt';
            else if (headerLower.includes('to')) icon = 'fa-map-pin';
            else if (headerLower.includes('odo') || headerLower.includes('start') || headerLower.includes('end')) icon = 'fa-tachometer-alt';
            else if (headerLower.includes('dist')) icon = 'fa-road';
            else if (headerLower.includes('coord')) icon = 'fa-globe';
            else if (headerLower.includes('email')) icon = 'fa-envelope';
            else if (headerLower.includes('notes')) icon = 'fa-sticky-note';
            else if (headerLower.includes('file')) icon = 'fa-file';
            else if (headerLower.includes('contact')) icon = 'fa-address-book';
            
            colHeader.innerHTML = `<i class="fas ${icon}"></i> ${header}`;
            headerRow.appendChild(colHeader);
        });
        
        if (tabName === 'MAIN') {
            const totalHeader = document.createElement('div');
            totalHeader.className = 'excel-col-header total-column';
            totalHeader.innerHTML = `<i class="fas fa-calculator"></i> Total (End-Start)`;
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
            
            const grandTotal = calculateGrandTotal(rowsToDisplay, headers);
            
            totalsContainer.innerHTML = `
                <span class="totals-label">
                    <i class="fas fa-calculator"></i> Grand Total (End - Start):
                </span>
                <span class="totals-value">${Math.round(grandTotal).toLocaleString()} km</span>
            `;
            
        } else {
            const allRows = tabData.data;
            const dateFilteredRows = filterVehicleRows(allRows);
            
            if (distIdx !== -1) {
                hiddenCount = allRows.filter(row => {
                    if (row[distIdx]) {
                        const distStr = row[distIdx].replace(',', '.').replace(/[^0-9.-]/g, '');
                        const distance = parseFloat(distStr) || 0;
                        return distance <= 0.10 && distance > 0;
                    }
                    return false;
                }).length;
            }
            
            rowsToDisplay = dateFilteredRows;
            const vehicleTotal = calculateVehicleTabTotal(rowsToDisplay, headers);
            
            if (filterStats) {
                filterStats.innerHTML = `Showing ${rowsToDisplay.length} of ${tabData.data.length} rows`;
                if (distIdx !== -1 && hiddenCount > 0) {
                    filterStats.innerHTML += `<span style="margin-left:15px; color:#1976d2;"><i class="fas fa-filter"></i> Hidden (Dist ≤ 0.10): ${hiddenCount} rows</span>`;
                }
            }
            
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
                
                if (tabName === 'MAIN' && (colIndex === startIdx || colIndex === endIdx)) {
                    cell.classList.add('odo-highlight');
                }
                
                if (tabName !== 'MAIN' && colIndex === distIdx) {
                    const distStr = row[colIndex] ? row[colIndex].replace(',', '.').replace(/[^0-9.-]/g, '') : '0';
                    const distance = parseFloat(distStr) || 0;
                    if (distance > 0.10) {
                        cell.style.fontWeight = '500';
                        cell.style.color = '#2e7d32';
                    }
                }
                
                cell.textContent = colIndex < row.length ? row[colIndex] : '-';
                excelRow.appendChild(cell);
            });
            
            if (tabName === 'MAIN') {
                const totalCell = document.createElement('div');
                totalCell.className = 'excel-cell total-cell';
                
                const rowTotal = calculateRowTotal(row, headers);
                totalCell.textContent = rowTotal !== null ? Math.round(rowTotal).toLocaleString() + ' km' : '-';
                excelRow.appendChild(totalCell);
            }
            
            excelGrid.appendChild(excelRow);
        });
        
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
        
        if (tabName === 'MAIN' && rowsToDisplay.length > 0) {
            const grandTotalRow = document.createElement('div');
            grandTotalRow.className = 'excel-row grand-total-row';
            
            headers.forEach(header => {
                if (!header || header.trim() === '') return;
                const cell = document.createElement('div');
                cell.className = 'excel-cell';
                cell.textContent = '';
                grandTotalRow.appendChild(cell);
            });
            
            const grandTotalCell = document.createElement('div');
            grandTotalCell.className = 'excel-cell grand-total-cell';
            const grandTotal = calculateGrandTotal(rowsToDisplay, headers);
            grandTotalCell.textContent = Math.round(grandTotal).toLocaleString() + ' km';
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
        
        // Update stats cards
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
                const variancePercent = mainTotal > 0 ? (variance / mainTotal * 100).toFixed(1) : '0';
                
                const card = document.createElement('div');
                card.className = 'vehicle-card';
                card.innerHTML = `
                    <div class="vehicle-title">
                        <i class="fas fa-truck"></i> ${vehicle}
                    </div>
                    <div class="vehicle-stats">
                        <div>
                            <div class="stat-label">MAIN</div>
                            <div class="stat-number main">${Math.round(mainTotal).toLocaleString()} km</div>
                        </div>
                        <div>
                            <div class="stat-label">TAB</div>
                            <div class="stat-number vehicle">${Math.round(vehicleTotal).toLocaleString()} km</div>
                        </div>
                    </div>
                    <div style="text-align: center;">
                        <span class="variance-badge ${variance >= 0 ? 'positive' : 'negative'}">
                            ${variance >= 0 ? '+' : ''}${Math.round(variance).toLocaleString()} km (${variancePercent}%)
                        </span>
                    </div>
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
                        label: `${vehicle} (TAB)`,
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
                const isMain = !dataset.borderDash || dataset.borderDash.length === 0;
                legend.innerHTML += `
                    <div class="legend-item">
                        <span class="legend-color" style="background-color: ${dataset.borderColor}"></span>
                        <span>${dataset.label} ${isMain ? '(MAIN)' : '(TAB)'}</span>
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
            headers.push(`${vehicle} MAIN`, `${vehicle} TAB`, 'Var');
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
        
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        
        const tabBtn = getElement(`tabBtn${view}`);
        if (tabBtn) tabBtn.classList.add('active');
        
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
            dataToExport = filterVehicleRows(tabData.data);
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
            html += '<th>Total (End - Start)</th>';
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
                const startIdx = findColumnIndex(headers, 'start');
                const endIdx = findColumnIndex(headers, 'end');
                
                if (startIdx !== -1 && endIdx !== -1 && row[startIdx] && row[endIdx]) {
                    const start = parseFloat(row[startIdx]) || 0;
                    const end = parseFloat(row[endIdx]) || 0;
                    const total = Math.max(0, end - start);
                    html += '<td>' + Math.round(total) + '</td>';
                } else {
                    html += '<td>-</td>';
                }
            } else {
                const distIdx = findColumnIndex(headers, 'dist');
                if (distIdx !== -1 && row[distIdx]) {
                    const distStr = row[distIdx].replace(',', '.').replace(/[^0-9.-]/g, '');
                    const distance = parseFloat(distStr) || 0;
                    html += '<td>' + distance.toFixed(2) + '</td>';
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
            html += `<th>${vehicle} MAIN (km)</th><th>${vehicle} TAB (km)</th><th>${vehicle} Variance</th>`;
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