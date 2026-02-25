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
        
        // Daily totals by source and vehicle
        let dailyTotals = {
            MAIN: {},      // { "BD78NGZN": { "2024-01-01": 150, ... }, ... }
            VEHICLE: {}     // { "BD78NGZN": { "2024-01-01": 145, ... }, ... }
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

        // Safe element getter
        function getElement(id) {
            const el = document.getElementById(id);
            if (!el) console.warn(`Element with id '${id}' not found`);
            return el;
        }

        // Helper: convert any date string to YYYY-MM-DD
        function normalizeDate(dateStr) {
            if (!dateStr) return '';
            let datePart = dateStr.split(' ')[0];
            let parts = datePart.split('/');
            if (parts.length === 3) {
                let month, day, year;
                if (parseInt(parts[0]) > 12) {
                    day = parts[0].padStart(2, '0');
                    month = parts[1].padStart(2, '0');
                    year = parts[2];
                } else {
                    month = parts[0].padStart(2, '0');
                    day = parts[1].padStart(2, '0');
                    year = parts[2];
                }
                if (year.length === 2) year = '20' + year;
                return `${year}-${month}-${day}`;
            }
            return datePart;
        }

        function isValidHeader(header) {
            if (!header || header.trim() === '') return false;
            const invalid = ['column', 'null', 'undefined', 'nan', '-', ''];
            return !invalid.includes(header.trim().toLowerCase());
        }

        function findColumnIndex(headers, searchTerm) {
            return headers.findIndex(header => 
                header && header.toLowerCase().includes(searchTerm.toLowerCase())
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
            
            updateDashboard();
        }

        // Apply custom date range
        function applyDateRange() {
            const startInput = getElement('startDate');
            const endInput = getElement('endDate');
            
            if (!startInput.value || !endInput.value) {
                showNotification('Please select both start and end dates', 'warning');
                return;
            }
            
            if (startInput.value > endInput.value) {
                showNotification('Start date must be before end date', 'warning');
                return;
            }
            
            currentDateRange.startDate = startInput.value;
            currentDateRange.endDate = endInput.value;
            
            updateDashboard();
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
            if (dates.length === 0) return 'No dates';
            const firstDate = new Date(dates[0]);
            const lastDate = new Date(dates[dates.length - 1]);
            
            const options = { month: 'short', day: 'numeric', year: 'numeric' };
            return `${firstDate.toLocaleDateString('en-US', options)} - ${lastDate.toLocaleDateString('en-US', options)}`;
        }

        async function loadAllSheets() {
            const loadingEl = getElement('loading');
            const errorEl = getElement('error');
            
            // loading is already visible by default (nicer start)
            
            if (errorEl) errorEl.classList.remove('active');
            
            try {
                allSheetData = {};
                
                for (const [tabName, tabInfo] of Object.entries(TABS)) {
                    try {
                        const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${tabInfo.gid}`;
                        const response = await fetch(url);
                        if (!response.ok) throw new Error(`Failed to load ${tabName}`);
                        
                        const csvText = await response.text();
                        const rows = parseCSV(csvText);
                        
                        const rawHeaders = rows.length > 0 ? rows[0] : [];
                        const validHeaders = rawHeaders.filter(h => isValidHeader(h));
                        const dataRows = rows.slice(1).filter(row => row.some(cell => cell && cell.trim() !== ''));
                        
                        allSheetData[tabName] = {
                            headers: validHeaders,
                            rawHeaders: rawHeaders,
                            data: dataRows,
                            type: tabInfo.type,
                            rawData: dataRows
                        };
                        
                        console.log(`✅ Loaded ${tabName}: ${validHeaders.length} columns, ${dataRows.length} rows`);
                        
                    } catch (error) {
                        console.error(`❌ Error loading ${tabName}:`, error);
                        allSheetData[tabName] = { headers: [], rawHeaders: [], data: [], type: tabInfo.type, rawData: [] };
                    }
                }
                
                // Process daily totals
                processDailyTotals();
                
                const footerTimestamp = getElement('footerTimestamp');
                if (footerTimestamp) footerTimestamp.textContent = new Date().toLocaleString();
                
                // Populate MAIN filter dropdowns
                populateMainFilters();
                
                // Set default to last 7 days
                setLast7Days();
                
                // Show dashboard
                switchToView('dashboard');
                showNotification('Data loaded successfully!', 'success');
                
                // Hide loading with fade
                if (loadingEl) loadingEl.classList.add('hidden');
                
            } catch (error) {
                if (errorEl) {
                    errorEl.textContent = 'Error loading sheets: ' + error.message;
                    errorEl.classList.add('active');
                }
                // Still hide loading but show error
                if (loadingEl) loadingEl.classList.add('hidden');
            }
        }

        function parseCSV(csvText) {
            const lines = csvText.split('\n');
            const rows = [];
            for (let line of lines) {
                if (line.trim() === '') continue;
                const row = [];
                let inQuote = false;
                let value = '';
                for (let char of line) {
                    if (char === '"') {
                        inQuote = !inQuote;
                    } else if (char === ',' && !inQuote) {
                        row.push(value);
                        value = '';
                    } else {
                        value += char;
                    }
                }
                row.push(value);
                rows.push(row);
            }
            return rows;
        }

        function processDailyTotals() {
            // Initialize daily totals for each vehicle
            VEHICLES.forEach(vehicle => {
                if (!dailyTotals.MAIN[vehicle]) dailyTotals.MAIN[vehicle] = {};
                if (!dailyTotals.VEHICLE[vehicle]) dailyTotals.VEHICLE[vehicle] = {};
            });
            
            // Process MAIN data
            const mainData = allSheetData['MAIN'];
            if (mainData && mainData.data) {
                const mainHeaders = mainData.rawHeaders;
                const mainRegIdx = findColumnIndex(mainHeaders, 'truck') !== -1 ? 
                    findColumnIndex(mainHeaders, 'truck') : findColumnIndex(mainHeaders, 'reg');
                const mainDateIdx = findColumnIndex(mainHeaders, 'date') !== -1 ? 
                    findColumnIndex(mainHeaders, 'date') : findColumnIndex(mainHeaders, 'timestamp');
                const mainStartIdx = findColumnIndex(mainHeaders, 'start');
                const mainEndIdx = findColumnIndex(mainHeaders, 'end');
                
                if (mainRegIdx !== -1 && mainDateIdx !== -1 && mainStartIdx !== -1 && mainEndIdx !== -1) {
                    mainData.data.forEach(row => {
                        const reg = row[mainRegIdx] || '';
                        const vehicle = extractRegistration(reg);
                        
                        if (vehicle && VEHICLES.includes(vehicle)) {
                            const date = normalizeDate(row[mainDateIdx] || '');
                            if (!date) return;
                            
                            const start = parseFloat(row[mainStartIdx]) || 0;
                            const end = parseFloat(row[mainEndIdx]) || 0;
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
                const vehicleData = allSheetData[vehicle];
                if (!vehicleData || !vehicleData.data) return;
                
                const vehicleHeaders = vehicleData.rawHeaders;
                const vehicleDateIdx = findColumnIndex(vehicleHeaders, 'stop time');
                const vehicleDistIdx = findColumnIndex(vehicleHeaders, 'dist');
                
                if (vehicleDateIdx === -1 || vehicleDistIdx === -1) return;
                
                vehicleData.data.forEach(row => {
                    const date = normalizeDate(row[vehicleDateIdx] || '');
                    if (!date) return;
                    
                    let distance = 0;
                    if (row[vehicleDistIdx]) {
                        const distStr = row[vehicleDistIdx].replace(',', '.').replace(/[^0-9.-]/g, '');
                        distance = parseFloat(distStr) || 0;
                    }
                    
                    if (!dailyTotals.VEHICLE[vehicle][date]) {
                        dailyTotals.VEHICLE[vehicle][date] = 0;
                    }
                    dailyTotals.VEHICLE[vehicle][date] += distance;
                });
            });
            
            console.log('Daily totals processed:', dailyTotals);
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
            
            // Re-render MAIN tab with filters
            if (currentView === 'MAIN') {
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
            
            if (currentView === 'MAIN') {
                displayTabData('MAIN');
            }
        }

        function filterMainRows(rows) {
            const headers = allSheetData['MAIN']?.rawHeaders || [];
            const driverIdx = findColumnIndex(headers, 'driver');
            const truckIdx = findColumnIndex(headers, 'truck') !== -1 ? 
                findColumnIndex(headers, 'truck') : findColumnIndex(headers, 'reg');
            const dateIdx = findColumnIndex(headers, 'date') !== -1 ? 
                findColumnIndex(headers, 'date') : findColumnIndex(headers, 'timestamp');
            
            return rows.filter(row => {
                // Driver filter
                if (mainFilters.driver !== 'all' && driverIdx !== -1) {
                    if (row[driverIdx] !== mainFilters.driver) return false;
                }
                
                // Truck filter
                if (mainFilters.truck !== 'all' && truckIdx !== -1) {
                    if (row[truckIdx] !== mainFilters.truck) return false;
                }
                
                // Date range filter
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

        function displayTabData(tabName) {
            const tabData = allSheetData[tabName];
            if (!tabData) {
                console.log(`No data for tab: ${tabName}`);
                return;
            }
            
            const columnHeaders = getElement('tabColumnHeaders');
            const excelGrid = getElement('tabExcelGrid');
            const tabTitle = getElement('currentTabTitle');
            const filterBar = getElement('mainFilterBar');
            const filterStats = getElement('mainFilterStats');
            
            if (!columnHeaders || !excelGrid) return;
            
            columnHeaders.innerHTML = '';
            excelGrid.innerHTML = '';
            
            if (tabTitle) tabTitle.textContent = tabName;
            
            // Show/hide filter bar based on tab
            if (filterBar) {
                filterBar.style.display = tabName === 'MAIN' ? 'flex' : 'none';
            }
            
            const headers = tabData.rawHeaders || [];
            
            // Create column headers
            headers.forEach((header, index) => {
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
                
                colHeader.innerHTML = `<i class="fas ${icon}"></i> ${header}`;
                columnHeaders.appendChild(colHeader);
            });
            
            // Get filtered data for MAIN tab
            let rowsToDisplay = tabData.data;
            let filteredCount = rowsToDisplay.length;
            
            if (tabName === 'MAIN') {
                rowsToDisplay = filterMainRows(tabData.data);
                filteredCount = rowsToDisplay.length;
                if (filterStats) {
                    filterStats.textContent = `Showing ${filteredCount} of ${tabData.data.length} rows`;
                }
            }
            
            // Create data rows
            rowsToDisplay.forEach((row, rowIndex) => {
                const excelRow = document.createElement('div');
                excelRow.className = 'excel-row';
                
                const rowNum = document.createElement('div');
                rowNum.className = 'excel-row-number';
                rowNum.textContent = rowIndex + 1;
                excelRow.appendChild(rowNum);
                
                headers.forEach((header, colIndex) => {
                    if (!header || header.trim() === '') return;
                    
                    const cell = document.createElement('div');
                    cell.className = 'excel-cell';
                    let value = colIndex < row.length ? row[colIndex] : '-';
                    
                    cell.textContent = value;
                    excelRow.appendChild(cell);
                });
                
                excelGrid.appendChild(excelRow);
            });
            
            // Update footer
            const footerVehicleCount = getElement('footerVehicleCount');
            if (footerVehicleCount) footerVehicleCount.textContent = rowsToDisplay.length;
        }

        function updateDashboard() {
            const selectedVehicle = getElement('vehicleFilter')?.value || 'all';
            
            // Get date range
            const dateRange = getDateRange();
            
            // Update date range display
            const dateRangeText = getElement('dateRangeText');
            const footerDateRange = getElement('footerDateRange');
            const formattedRange = formatDateRange(dateRange);
            if (dateRangeText) dateRangeText.textContent = formattedRange;
            if (footerDateRange) footerDateRange.textContent = formattedRange;
            
            // Get vehicles to display
            const vehiclesToShow = selectedVehicle === 'all' ? VEHICLES : [selectedVehicle];
            
            // Prepare chart data
            const chartData = prepareDateRangeData(vehiclesToShow, dateRange);
            updateChart(chartData);
            
            // Update vehicle stats cards
            updateVehicleStats(vehiclesToShow, dateRange);
            
            // Update comparison table
            updateComparisonTable(vehiclesToShow, dateRange);
            
            // Update footer
            const footerVehicleCount = getElement('footerVehicleCount');
            if (footerVehicleCount) footerVehicleCount.textContent = vehiclesToShow.length;
        }

        function prepareDateRangeData(vehicles, dateRange) {
            const result = {
                labels: dateRange.map(date => {
                    const d = new Date(date);
                    return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
                }),
                datasets: []
            };
            
            const colors = ['#4da6ff', '#ff9f4a', '#6c5ce7', '#00b894', '#e17055', '#0984e3'];
            let colorIndex = 0;
            
            vehicles.forEach(vehicle => {
                // Prepare data arrays with zeros for missing dates
                const mainData = dateRange.map(date => {
                    return Math.round(dailyTotals.MAIN[vehicle]?.[date] || 0);
                });
                
                const vehicleData = dateRange.map(date => {
                    return Math.round(dailyTotals.VEHICLE[vehicle]?.[date] || 0);
                });
                
                // MAIN dataset (solid line)
                result.datasets.push({
                    label: `${vehicle} (MAIN)`,
                    data: mainData,
                    borderColor: colors[colorIndex],
                    backgroundColor: 'transparent',
                    borderWidth: 1.2,
                    borderDash: [],
                    tension: 0.2,
                    fill: false,
                    pointRadius: 2,
                    pointHoverRadius: 3
                });
                
                // VEHICLE dataset (dashed line)
                result.datasets.push({
                    label: `${vehicle} (TAB)`,
                    data: vehicleData,
                    borderColor: colors[colorIndex],
                    backgroundColor: 'transparent',
                    borderWidth: 1.2,
                    borderDash: [4, 3],
                    tension: 0.2,
                    fill: false,
                    pointRadius: 2,
                    pointHoverRadius: 3
                });
                
                colorIndex = (colorIndex + 1) % colors.length;
            });
            
            return result;
        }

        function updateChart(chartData) {
            if (!chartData) return;
            
            const canvas = getElement('distanceChart');
            if (!canvas) return;
            
            const ctx = canvas.getContext('2d');
            
            if (distanceChart) {
                distanceChart.destroy();
            }
            
            // Update legend
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
                distanceChart = new Chart(ctx, {
                    type: 'line',
                    data: {
                        labels: chartData.labels,
                        datasets: chartData.datasets
                    },
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
                            x: { 
                                ticks: { maxRotation: 45, minRotation: 45, font: { size: 8 } },
                                grid: { display: true, color: 'rgba(128,128,128,0.1)' }
                            },
                            y: {
                                beginAtZero: true,
                                title: { display: true, text: 'km', font: { size: 8 } },
                                ticks: { callback: v => v.toLocaleString(), font: { size: 8 } },
                                grid: { display: true, color: 'rgba(128,128,128,0.1)' }
                            }
                        }
                    }
                });
            } catch (e) {
                console.error('Chart error:', e);
            }
        }

        function updateVehicleStats(vehicles, dateRange) {
            const statsContainer = getElement('vehicleStatsCards');
            if (!statsContainer) return;
            
            statsContainer.innerHTML = '';
            
            vehicles.forEach(vehicle => {
                // Calculate totals for the date range
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

        function updateComparisonTable(vehicles, dateRange) {
            const headersContainer = getElement('comparisonHeaders');
            const rowsContainer = getElement('comparisonRows');
            
            if (!headersContainer || !rowsContainer) return;
            
            headersContainer.innerHTML = '';
            rowsContainer.innerHTML = '';
            
            // Create headers
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
            
            // Map vehicle to CSS class for text color
            const vehicleClassMap = {
                'BD78NGZN': 'vehicle-bd78',
                'CS44GHNZ': 'vehicle-cs44',
                'DG28ZLZN': 'vehicle-dg28'
            };
            
            // Create rows for each day in date range
            dateRange.forEach((date, index) => {
                const row = document.createElement('div');
                row.className = 'comp-row';
                
                // Check if this day has any data
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
                
                // Row number
                const rowNum = document.createElement('div');
                rowNum.className = 'row-number';
                rowNum.textContent = index + 1;
                row.appendChild(rowNum);
                
                // Date
                const dateCell = document.createElement('div');
                dateCell.className = 'comp-cell';
                const displayDate = new Date(date).toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
                dateCell.textContent = displayDate;
                if (!hasData) dateCell.classList.add('zero-value');
                row.appendChild(dateCell);
                
                // Data for each vehicle
                vehicles.forEach(vehicle => {
                    const mainVal = dailyTotals.MAIN[vehicle]?.[date] || 0;
                    const vehicleVal = dailyTotals.VEHICLE[vehicle]?.[date] || 0;
                    const variance = mainVal - vehicleVal;
                    const vehicleClass = vehicleClassMap[vehicle] || '';
                    
                    // MAIN value - color by vehicle
                    const mainCell = document.createElement('div');
                    mainCell.className = `comp-cell ${vehicleClass}`;
                    mainCell.textContent = mainVal > 0 ? Math.round(mainVal).toLocaleString() : '-';
                    if (mainVal === 0) mainCell.classList.add('zero-value');
                    row.appendChild(mainCell);
                    
                    // Vehicle value - color by vehicle
                    const vehicleCell = document.createElement('div');
                    vehicleCell.className = `comp-cell ${vehicleClass}`;
                    vehicleCell.textContent = vehicleVal > 0 ? Math.round(vehicleVal).toLocaleString() : '-';
                    if (vehicleVal === 0) vehicleCell.classList.add('zero-value');
                    row.appendChild(vehicleCell);
                    
                    // Variance - also color by vehicle (with background highlight)
                    const varCell = document.createElement('div');
                    varCell.className = `comp-cell ${vehicleClass}`;
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
            
            // Update tab buttons
            document.querySelectorAll('.tab-btn').forEach(btn => {
                btn.classList.remove('active');
            });
            
            const tabBtn = getElement(`tabBtn${view}`);
            if (tabBtn) tabBtn.classList.add('active');
            
            // Update view display
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
                updateDashboard();
            } else {
                if (dashboardView) dashboardView.style.display = 'none';
                if (tableView) {
                    tableView.style.display = 'flex';
                    displayTabData(view);
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
            
            // For MAIN tab, export filtered data
            let dataToExport = tabData.data;
            if (currentView === 'MAIN') {
                dataToExport = filterMainRows(tabData.data);
            }
            
            const headers = tabData.rawHeaders || [];
            let csvContent = headers.join(',') + '\n';
            
            dataToExport.forEach(row => {
                csvContent += row.map(cell => `"${cell}"`).join(',') + '\n';
            });
            
            const blob = new Blob([csvContent], { type: 'text/csv' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.download = `${currentView}-${new Date().toISOString().slice(0,10)}.csv`;
            a.href = url;
            a.click();
            URL.revokeObjectURL(url);
            
            showNotification('Export successful!', 'success');
        }

        function exportComparisonData() {
            const selectedVehicle = getElement('vehicleFilter')?.value || 'all';
            const vehicles = selectedVehicle === 'all' ? VEHICLES : [selectedVehicle];
            const dateRange = getDateRange();
            
            // Create CSV
            let csv = 'Date';
            vehicles.forEach(v => {
                csv += `,${v} MAIN (km),${v} TAB (km),${v} Variance`;
            });
            csv += '\n';
            
            dateRange.forEach(date => {
                let row = new Date(date).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
                vehicles.forEach(vehicle => {
                    const mainVal = dailyTotals.MAIN[vehicle]?.[date] || 0;
                    const vehicleVal = dailyTotals.VEHICLE[vehicle]?.[date] || 0;
                    const variance = mainVal - vehicleVal;
                    
                    row += `,${Math.round(mainVal)},${Math.round(vehicleVal)},${Math.round(variance)}`;
                });
                csv += row + '\n';
            });
            
            const blob = new Blob([csv], { type: 'text/csv' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.download = `date-range-${selectedVehicle}-${new Date().toISOString().slice(0,10)}.csv`;
            a.href = url;
            a.click();
            URL.revokeObjectURL(url);
            
            showNotification('Export successful!', 'success');
        }

        function showNotification(message, type) {
            const notification = document.createElement('div');
            notification.className = 'excel-notification';
            notification.innerHTML = `<i class="fas ${type === 'success' ? 'fa-check-circle' : 'fa-info-circle'}"></i> ${message}`;
            document.body.appendChild(notification);
            
            setTimeout(() => {
                notification.style.opacity = '0';
                setTimeout(() => notification.remove(), 300);
            }, 3000);
        }

        // Initialize when DOM is ready
        document.addEventListener('DOMContentLoaded', function() {
            console.log('DOM ready, loading sheets...');
            // Show nice loading by default (already visible)
            setTimeout(loadAllSheets, 100);
        });
