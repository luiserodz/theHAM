// Global variables
let csvData = [];
let charts = {};

// Load logo image as data URL
// If the logo cannot be loaded (e.g., running from the local file system),
// return null so PDF generation can continue without it.
async function loadLogo() {
    try {
        const response = await fetch('primary-branding-asset.png');
        if (!response.ok) throw new Error('Network response was not ok');
        const blob = await response.blob();
        return new Promise((resolve) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result);
            reader.readAsDataURL(blob);
        });
    } catch (error) {
        console.warn('Logo could not be loaded:', error);
        return null;
    }
}

// DOM elements
const csvFile = document.getElementById('csvFile');
const chooseFileBtn = document.getElementById('chooseFileBtn');
const uploadSection = document.getElementById('uploadSection');
const dataPreview = document.getElementById('dataPreview');
const reportOptions = document.getElementById('reportOptions');
const chartsPreview = document.getElementById('chartsPreview');
const statusSection = document.getElementById('statusSection');
const downloadSection = document.getElementById('downloadSection');
const generateReportBtn = document.getElementById('generateReport');
const downloadLink = document.getElementById('downloadLink');

// Initialize on page load
document.addEventListener('DOMContentLoaded', function() {
    console.log('Windows Update Report Generator initializing...');
    setupEventListeners();
    loadSavedPreferences();
    downloadLink.classList.add('disabled');
    console.log('Initialization complete');
});

function setupEventListeners() {
    // File upload button
    chooseFileBtn.addEventListener('click', () => {
        csvFile.click();
    });
    
    // File input change
    csvFile.addEventListener('change', handleFileUpload);
    
    // Drag and drop
    uploadSection.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadSection.classList.add('dragover');
    });
    
    uploadSection.addEventListener('dragleave', () => {
        uploadSection.classList.remove('dragover');
    });
    
    uploadSection.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadSection.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 0 && files[0].name.endsWith('.csv')) {
            csvFile.files = files;
            handleFileUpload();
        }
    });
    
    // Sample CSV download
    document.getElementById('downloadSample').addEventListener('click', downloadSampleCSV);
    
    // Generate report button
    generateReportBtn.addEventListener('click', generatePDFReport);
    
    // Save form data on change
    document.querySelectorAll('#reportOptions input, #reportOptions select, #reportOptions textarea').forEach(element => {
        element.addEventListener('change', saveFormData);
    });

}

function handleFileUpload() {
    const file = csvFile.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            parseCSVData(e.target.result);
            showDataPreview();
            showReportOptions();
            generateCharts();
            generateReportBtn.disabled = false;
            showAlert('File uploaded successfully!', 'success');
        } catch (error) {
            showAlert('Error parsing CSV file: ' + error.message, 'error');
            console.error('Parse error:', error);
        }
    };
    reader.onerror = function() {
        showAlert('Error reading file', 'error');
    };
    reader.readAsText(file);
}

function parseCSVData(csvText) {
    const lines = csvText.trim().split('\n');
    if (lines.length < 2) {
        throw new Error('CSV file must contain headers and at least one data row');
    }

    // Parse headers and normalize common names (case/spacing differences)
    const rawHeaders = parseCSVLine(lines[0]);
    const headerMap = {};
    rawHeaders.forEach(h => {
        const key = h.trim().toLowerCase().replace(/\s+/g, '');
        switch (key) {
            case 'updatetitle':
            case 'title':
                headerMap[h] = 'Update Name';
                break;
            case 'securityseverity':
            case 'severity':
                headerMap[h] = 'Security Severity';
                break;
            case 'releasedate':
            case 'releasedateutc':
                headerMap[h] = 'Release Date';
                break;
            case 'updatesmissing':
            case 'missingupdates':
                headerMap[h] = 'Updates Missing';
                break;
            case 'recentlydeployed':
            case 'deployed':
                headerMap[h] = 'Recently Deployed';
                break;
            case 'status':
            case 'updatestatus':
            case 'installstatus':
            case 'updatestatusofinstallation':
                headerMap[h] = 'Status';
                break;
            default:
                headerMap[h] = h.trim();
        }
    });

    // Parse data rows
    csvData = [];
    for (let i = 1; i < lines.length; i++) {
        const values = parseCSVLine(lines[i]);
        if (values.length === rawHeaders.length) {
            const row = {};
            rawHeaders.forEach((header, index) => {
                row[headerMap[header]] = values[index] || '';
            });
            if (row['Update Name']) { // Only add rows with update names
                csvData.push(row);
            }
        }
    }

    if (csvData.length === 0) {
        throw new Error('No valid data found in CSV');
    }

    console.log('Parsed CSV data:', csvData.length, 'rows');
}

function parseCSVLine(line) {
    const result = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        
        if (char === '"') {
            inQuotes = !inQuotes;
        } else if (char === ',' && !inQuotes) {
            result.push(current.trim());
            current = '';
        } else {
            current += char;
        }
    }
    
    result.push(current.trim());
    return result;
}

// ============= ENHANCED EXECUTIVE SUMMARY FUNCTIONS =============

// Calculate days since release
function calculateDaysOld(releaseDateStr) {
    if (!releaseDateStr) return 0;
    try {
        const releaseDate = new Date(releaseDateStr.replace(/-/g, '/'));
        if (isNaN(releaseDate.getTime())) return 0;
        const daysSince = Math.floor((new Date() - releaseDate) / (1000 * 60 * 60 * 24));
        return Math.max(0, daysSince);
    } catch (e) {
        return 0;
    }
}

// Calculate update age distribution
function calculateUpdateAges(csvData) {
    const ageGroups = {
        '0-7 days': 0,
        '8-14 days': 0,
        '15-30 days': 0,
        '31-60 days': 0,
        '>60 days': 0
    };
    
    csvData.forEach(row => {
        const daysSince = calculateDaysOld(row['Release Date']);
        
        if (daysSince <= 7) ageGroups['0-7 days']++;
        else if (daysSince <= 14) ageGroups['8-14 days']++;
        else if (daysSince <= 30) ageGroups['15-30 days']++;
        else if (daysSince <= 60) ageGroups['31-60 days']++;
        else ageGroups['>60 days']++;
    });
    
    return ageGroups;
}

function getSortedData() {
    const severityOrder = { 'critical': 0, 'important': 1, 'moderate': 2, 'low': 3 };
    return csvData.slice().sort((a, b) => {
        const aSev = severityOrder[a['Security Severity']?.toLowerCase()] ?? 4;
        const bSev = severityOrder[b['Security Severity']?.toLowerCase()] ?? 4;
        if (aSev !== bSev) return aSev - bSev;
        const aMiss = parseInt(a['Updates Missing']?.toString().replace(/\D/g, '') || '0');
        const bMiss = parseInt(b['Updates Missing']?.toString().replace(/\D/g, '') || '0');
        return bMiss - aMiss;
    });
}

function getStatusWithReboot(row) {
    let status = row['Status'] || '';
    const searchSpace = Object.values(row).join(' ').toLowerCase();
    if (searchSpace.includes('installation has been initiated')) {
        status = status ? `${status} - Pending Reboot` : 'Pending Reboot';
    }
    return status;
}

// Get top deployment gaps
function getTopDeploymentGaps(csvData) {
    return csvData
        .filter(row => parseInt(row['Updates Missing']?.toString().replace(/\D/g, '') || '0') > 0)
        .map(row => ({
            updateName: row['Update Name'],
            version: row['Version'],
            missing: parseInt(row['Updates Missing']?.toString().replace(/\D/g, '') || '0'),
            severity: row['Security Severity'] || 'Unspecified',
            releaseDate: row['Release Date'],
            daysOld: calculateDaysOld(row['Release Date'])
        }))
        .sort((a, b) => {
            // Sort by severity first, then by missing count
            const severityOrder = { 'critical': 0, 'important': 1, 'moderate': 2, 'low': 3 };
            const aSev = severityOrder[a.severity.toLowerCase()] ?? 4;
            const bSev = severityOrder[b.severity.toLowerCase()] ?? 4;
            
            if (aSev !== bSev) return aSev - bSev;
            return b.missing - a.missing;
        })
        .slice(0, 5);
}

// Get top critical updates
function getTopCriticalUpdates(csvData) {
    return csvData
        .filter(row => row['Security Severity']?.toLowerCase().includes('critical'))
        .map(row => ({
            updateName: row['Update Name'],
            releaseDate: row['Release Date'],
            missing: parseInt(row['Updates Missing']?.toString().replace(/\D/g, '') || '0')
        }))
        .sort((a, b) => b.missing - a.missing)
        .slice(0, 5);
}

// Calculate risk scores
function calculateRiskScores(csvData) {
    const stats = generateStatistics();
    const critical = csvData.filter(r => 
        r['Security Severity']?.toLowerCase().includes('critical')
    );
    
    // Security Risk (based on critical updates and age)
    const oldCritical = critical.filter(r => 
        calculateDaysOld(r['Release Date']) > 30
    ).length;
    const securityRisk = Math.min(10, (critical.length * 0.5) + (oldCritical * 2));
    
    // Operational Risk (based on total missing updates)
    const operationalRisk = Math.min(10, stats.totalMissing / 100);
    
    // Compliance Risk (based on compliance rate)
    const complianceRisk = (100 - stats.complianceRate) / 10;
    
    return {
        security: Math.round(securityRisk),
        operational: Math.round(operationalRisk),
        compliance: Math.round(complianceRisk)
    };
}

// Calculate success metrics
function calculateSuccessMetrics(csvData) {
    const stats = generateStatistics();
    
    const fullyDeployed = csvData.filter(row => 
        parseInt(row['Updates Missing']?.toString().replace(/\D/g, '') || '0') === 0
    ).length;
    
    const defenderUpdates = csvData.filter(row => 
        row['Update Name'].toLowerCase().includes('defender')
    );
    const defenderCompliance = defenderUpdates.length > 0 ? 
        (defenderUpdates.filter(r => parseInt(r['Updates Missing']?.toString().replace(/\D/g, '') || '0') === 0).length / 
         defenderUpdates.length) * 100 : 100;
    
    return {
        totalDeployed: stats.totalDeployed,
        fullyDeployedUpdates: fullyDeployed,
        defenderCompliance: Math.round(defenderCompliance),
        perfectDeployments: csvData.filter(row => {
            const deployed = parseInt(row['Recently Deployed']?.toString().replace(/\D/g, '') || '0');
            return deployed > 0 && parseInt(row['Updates Missing']?.toString().replace(/\D/g, '') || '0') === 0;
        }).length
    };
}

function getExecutiveSummaryData() {
    const stats = generateStatistics();
    const totalDevices = stats.totalDeployed + stats.totalMissing;
    const criticalUpdates = csvData.filter(r =>
        r['Security Severity']?.toLowerCase().includes('critical')
    ).sort((a, b) => {
        const aMiss = parseInt(a['Updates Missing']?.toString().replace(/\D/g, '') || '0');
        const bMiss = parseInt(b['Updates Missing']?.toString().replace(/\D/g, '') || '0');
        return bMiss - aMiss;
    }).slice(0, 2);

    const topCritical = criticalUpdates.map(u => `${u['Update Name']} (${u['Updates Missing']})`);

    return {
        complianceRate: stats.complianceRate,
        totalDevices,
        topCritical
    };
}

function showDataPreview() {
    if (csvData.length === 0) return;
    
    // Generate statistics
    const stats = generateStatistics();
    displayStatistics(stats);
    updateEnhancedPreview();
    
    // Show first 10 rows in table
    const table = document.getElementById('previewTable');
    const sorted = getSortedData();
    const headers = Object.keys(sorted[0]);
    
    let tableHTML = '<thead><tr>';
    headers.forEach(h => {
        tableHTML += `<th>${h}</th>`;
    });
    tableHTML += '</tr></thead><tbody>';
    
    sorted.slice(0, 10).forEach(row => {
        const pending = parseInt(row['Updates Missing']?.toString().replace(/\D/g, '') || '0');
        const daysOld = calculateDaysOld(row['Release Date']);
        let rowClass = '';
        if (pending > 10) rowClass = 'row-high-missing';
        else if (daysOld > 30) rowClass = 'row-old';

        tableHTML += `<tr class="${rowClass}">`;
        headers.forEach(h => {
            let value = row[h];
            if (h === 'Status') value = getStatusWithReboot(row);
            const className = getSeverityClass(value);
            tableHTML += `<td class="${className}">${value}</td>`;
        });
        tableHTML += '</tr>';
    });
    
    tableHTML += '</tbody>';
    table.innerHTML = tableHTML;
    
    dataPreview.classList.add('show');
}

function updateEnhancedPreview() {
    // Refresh age chart if present
    if (charts.age) {
        charts.age.destroy();
    }
    const ageCtx = document.getElementById('ageChart');
    if (ageCtx) {
        charts.age = generateAgeChart();
    }
}

function generateStatistics() {
    const totalUpdates = csvData.length;
    
    // Parse numbers correctly, removing any non-numeric characters
    const totalMissing = csvData.reduce((sum, row) => {
        const missing = parseInt(row['Updates Missing']?.toString().replace(/\D/g, '') || '0');
        return sum + (isNaN(missing) ? 0 : missing);
    }, 0);
    
    const totalDeployed = csvData.reduce((sum, row) => {
        const deployed = parseInt(row['Recently Deployed']?.toString().replace(/\D/g, '') || '0');
        return sum + (isNaN(deployed) ? 0 : deployed);
    }, 0);
    
    const severityCounts = csvData.reduce((acc, row) => {
        const severity = row['Security Severity'] || 'None';
        acc[severity] = (acc[severity] || 0) + 1;
        return acc;
    }, {});
    
    const criticalUpdates = severityCounts['Critical'] || 0;
    const importantUpdates = severityCounts['Important'] || 0;
    
    // Calculate compliance rate
    const totalDeploymentTargets = totalDeployed + totalMissing;
    const complianceRate = totalDeploymentTargets > 0 ? Math.round((totalDeployed / totalDeploymentTargets) * 100) : 100;
    
    return {
        totalUpdates,
        totalMissing,
        totalDeployed,
        criticalUpdates,
        importantUpdates,
        severityCounts,
        complianceRate
    };
}

function displayStatistics(stats) {
    const statsContainer = document.getElementById('dataStats');
    statsContainer.innerHTML = `
        <div class="stat-card">
            <div class="stat-number">${stats.totalUpdates}</div>
            <div class="stat-label">Total Updates</div>
        </div>
        <div class="stat-card">
            <div class="stat-number severity-critical">${stats.criticalUpdates}</div>
            <div class="stat-label">Critical Updates</div>
        </div>
        <div class="stat-card">
            <div class="stat-number severity-important">${stats.importantUpdates}</div>
            <div class="stat-label">Important Updates</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">${stats.totalMissing}</div>
            <div class="stat-label">Missing Deployments</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">${stats.complianceRate}%</div>
            <div class="stat-label">Compliance Rate</div>
        </div>
    `;
}

function getSeverityClass(value) {
    if (!value) return '';
    const severity = value.toString().toLowerCase();
    if (severity.includes('critical')) return 'severity-critical';
    if (severity.includes('important')) return 'severity-important';
    if (severity.includes('moderate')) return 'severity-moderate';
    if (severity.includes('low')) return 'severity-low';
    return '';
}

function showReportOptions() {
    reportOptions.classList.add('show');
}

function generateCharts() {
    if (typeof Chart === 'undefined') {
        console.warn('Chart.js not loaded; skipping chart generation');
        showAlert('Charts could not be loaded. PDF will exclude chart visuals.', 'info');
        chartsPreview.classList.remove('show');
        return;
    }

    generateAgeChart();
    generateSeverityChart();
    generateDeploymentChart();
    generateTrendChart();
    chartsPreview.classList.add('show');
}

function generateAgeChart() {
    if (typeof Chart === 'undefined') return;
    const ctx = document.getElementById('ageChart').getContext('2d');
    const ages = calculateUpdateAges(csvData);
    const labels = Object.keys(ages);
    const data = Object.values(ages);

    if (charts.age) {
        charts.age.destroy();
    }

    charts.age = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                data: data,
                backgroundColor: ['#28a745','#20c997','#ffc107','#fd7e14','#dc3545']
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: {
                    beginAtZero: true,
                    ticks: {
                        font: { weight: 'bold' }
                    }
                },
                y: {
                    ticks: {
                        font: { weight: 'bold' }
                    }
                }
            },
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: ctx => `${ctx.parsed.x} updates`
                    }
                }
            }
        }
    });
    return charts.age;
}

function generateSeverityChart() {
    if (typeof Chart === 'undefined') return;
    const ctx = document.getElementById('severityChart').getContext('2d');
    
    const severityCounts = csvData.reduce((acc, row) => {
        const severity = row['Security Severity'] || 'None';
        acc[severity] = (acc[severity] || 0) + 1;
        return acc;
    }, {});
    
    const colors = {
        'Critical': '#dc3545',
        'Important': '#fd7e14',
        'Moderate': '#ffc107',
        'Low': '#28a745',
        'None': '#6c757d'
    };
    
    if (charts.severity) {
        charts.severity.destroy();
    }
    
    // Calculate percentages for display
    const total = Object.values(severityCounts).reduce((a, b) => a + b, 0);
    const dataWithPercentages = Object.keys(severityCounts).map(key => {
        const value = severityCounts[key];
        const percentage = ((value / total) * 100).toFixed(1);
        return `${key} (${percentage}%)`;
    });
    
    charts.severity = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: dataWithPercentages,
            datasets: [{
                data: Object.values(severityCounts),
                backgroundColor: Object.keys(severityCounts).map(s => colors[s] || '#6c757d'),
                borderWidth: 2,
                borderColor: '#fff'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'right',
                    labels: {
                        font: {
                            size: 14,
                        },
                        boxWidth: 25,
                        padding: 15,
                        generateLabels: function(chart) {
                            const data = chart.data;
                            return data.labels.map((label, i) => ({
                                text: label,
                                fillStyle: data.datasets[0].backgroundColor[i],
                                hidden: false,
                                index: i
                            }));
                        }
                    }
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const label = context.label || '';
                            const value = context.parsed;
                            return `${label}: ${value} updates`;
                        }
                    }
                }
            }
        }
    });
}

function generateDeploymentChart() {
    if (typeof Chart === 'undefined') return;
    const ctx = document.getElementById('deploymentChart').getContext('2d');
    
    const deploymentData = csvData.map(row => ({
        name: row['Update Name'],
        deployed: parseInt(row['Recently Deployed']?.toString().replace(/\D/g, '') || '0'),
        missing: parseInt(row['Updates Missing']?.toString().replace(/\D/g, '') || '0')
    }))
    .filter(d => d.deployed > 0 || d.missing > 0)
    .sort((a, b) => (b.deployed + b.missing) - (a.deployed + a.missing))
    .slice(0, 10);
    
    if (charts.deployment) {
        charts.deployment.destroy();
    }
    
    charts.deployment = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: deploymentData.map(d => {
                // Better truncation with full name in tooltip
                if (d.name.length > 50) {
                    return d.name.substring(0, 47) + '...';
                }
                return d.name;
            }),
            datasets: [
                {
                    label: 'Deployed',
                    data: deploymentData.map(d => d.deployed),
                    backgroundColor: '#28a745',
                    borderWidth: 1,
                    borderColor: '#1e7e34'
                },
                {
                    label: 'Pending Updates',
                    data: deploymentData.map(d => d.missing),
                    backgroundColor: '#dc3545',
                    borderWidth: 1,
                    borderColor: '#bd2130'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: {
                    ticks: {
                        font: {
                            size: 11
                        },
                        autoSkip: false,
                        maxRotation: 45,
                        minRotation: 45
                    }
                },
                y: {
                    beginAtZero: true,
                    ticks: {
                        font: {
                            size: 14,
                        }
                    }
                }
            },
            plugins: {
                legend: {
                    position: 'top',
                    labels: {
                        font: {
                            size: 14,
                        },
                        padding: 20
                    }
                },
                tooltip: {
                    callbacks: {
                        title: (context) => {
                            // Show full name in tooltip
                            return [deploymentData[context[0].dataIndex].name];
                        },
                        label: (context) => {
                            const label = context.dataset.label || '';
                            const value = context.parsed.y;
                            const total = deploymentData[context.dataIndex].deployed + deploymentData[context.dataIndex].missing;
                            const percentage = total > 0 ? ((value / total) * 100).toFixed(1) : 0;
                            return `${label}: ${value} systems (${percentage}%)`;
                        }
                    }
                }
            }
        }
    });
}

function generateTrendChart() {
    if (typeof Chart === 'undefined') return;
    const ctx = document.getElementById('trendChart').getContext('2d');
    
    // Group by month from Release Date
    const monthCounts = {};
    
    csvData.forEach(row => {
        const dateStr = row['Release Date'];
        if (dateStr) {
            try {
                const date = new Date(dateStr.replace(/-/g, '/')); // More reliable parsing
                if (!isNaN(date.getTime())) {
                    const monthKey = date.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
                    monthCounts[monthKey] = (monthCounts[monthKey] || 0) + 1;
                }
            } catch (e) {
                console.warn(`Could not parse date: ${dateStr}`);
            }
        }
    });
    
    if (charts.trend) {
        charts.trend.destroy();
    }
    
    const sortedMonths = Object.keys(monthCounts).sort((a, b) => new Date(a) - new Date(b));
    
    charts.trend = new Chart(ctx, {
        type: 'line',
        data: {
            labels: sortedMonths,
            datasets: [{
                label: 'Updates Released',
                data: sortedMonths.map(month => monthCounts[month]),
                borderColor: '#0078d4',
                backgroundColor: 'rgba(0, 120, 212, 0.1)',
                fill: true,
                tension: 0.4,
                pointRadius: 5,
                pointHoverRadius: 7,
                pointBackgroundColor: '#0078d4',
                pointBorderColor: '#fff',
                pointBorderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                 x: {
                    ticks: {
                        font: {
                            size: 14,
                        }
                    }
                },
                y: {
                    beginAtZero: true,
                    ticks: {
                        font: {
                            size: 14,
                        },
                        stepSize: 1
                    }
                }
            },
            plugins: {
                legend: {
                    position: 'top',
                    labels: {
                        font: {
                            size: 14,
                        },
                        padding: 20
                    }
                },
                tooltip: {
                    callbacks: {
                        label: (context) => {
                            return `${context.dataset.label}: ${context.parsed.y} updates`;
                        }
                    }
                }
            }
        }
    });
}

async function generatePDFReport() {
    statusSection.classList.add('show');
    // Hide any previous download section before starting a new run
    downloadSection.classList.remove('show');
    downloadLink.classList.add('disabled');
    downloadLink.removeAttribute('href');
    downloadLink.removeAttribute('download');
    statusSection.scrollIntoView({ behavior: 'smooth' });
    updateProgress(0, 'Initializing report generation...');
    
    try {
        // Initialize jsPDF from global
        const jsPDFLib = window.jspdf ? window.jspdf.jsPDF : window.jsPDF;
        if (!jsPDFLib) {
            throw new Error('jsPDF library failed to load');
        }
        const pdf = new jsPDFLib('p', 'mm', 'a4');
        
        updateProgress(10, 'Creating report structure...');
        
        // Get form values
        const reportTitle = document.getElementById('reportTitle').value;
        const orgName = document.getElementById('organizationName').value;
        const reportPeriod = document.getElementById('reportPeriod').value;
        const additionalNotes = document.getElementById('additionalNotes').value;

        const logo = await loadLogo();

        // Generate report content
        await createPDFContent(pdf, {
            title: reportTitle,
            organization: orgName,
            period: reportPeriod,
            notes: additionalNotes,
            logo: logo
        });
        
        updateProgress(100, 'Report generation complete!');
        
        // Create download link
        const pdfBlob = pdf.output('blob');
        const url = URL.createObjectURL(pdfBlob);
        downloadLink.href = url;
        downloadLink.download = `${reportTitle.replace(/[^a-z0-9]/gi, '_')}_${new Date().toISOString().slice(0, 10)}.pdf`;
        downloadLink.classList.remove('disabled');

        // Show download section and bring it into view
        downloadSection.classList.add('show');
        statusSection.classList.remove('show');
        downloadLink.scrollIntoView({ behavior: 'smooth' });

        // Focus download link for accessibility
        downloadLink.focus();
        
    } catch (error) {
        console.error('Error generating PDF:', error);
        showAlert('Error generating PDF: ' + error.message, 'error');
        statusSection.classList.remove('show');
    }
}

async function createPDFContent(pdf, config) {
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const pageMargin = 20; // uniform left/right margin
    const usableWidth = pageWidth - pageMargin * 2;
    let currentPage = 1;

    // Dimensions used when embedding chart images
    const chartWidth = pageWidth - pageMargin * 2;
    const chartHeight = 60;
    
    // Standardized font sizes
    const FONT_SIZES = {
        title: 24,
        subtitle: 14,
        sectionTitle: 14,
        heading: 12,
        body: 10,
        small: 9,
        caption: 8
    };
    
    // Standard spacing values
    const lineHeight = 6;
    const paragraphSpacing = 4;
    const sectionSpacing = 10;
    
    // --- TITLE PAGE ---
    updateProgress(20, 'Creating title page...');

    // Collect overall statistics for the report
    const reportStats = generateStatistics();
    const updateAges = calculateUpdateAges(csvData);
    const topCriticalUpdates = getTopCriticalUpdates(csvData);

    // Header
    pdf.setFillColor(0, 120, 212);
    const headerHeight = 50; // Reduced from 60
    pdf.rect(0, 0, pageWidth, headerHeight, 'F');

    let headerY = 8; // Reduced from 10
    if (config.logo) {
        const props = pdf.getImageProperties(config.logo);
        const logoWidth = 20; // Reduced from 25
        const logoHeight = (props.height / props.width) * logoWidth;
        const logoX = (pageWidth - logoWidth) / 2;
        pdf.addImage(config.logo, 'PNG', logoX, headerY, logoWidth, logoHeight);
        // Increase spacing below the logo so the title aligns more evenly with the organization name
        headerY += logoHeight + 8;
    }

    pdf.setTextColor(255, 255, 255);
    pdf.setFontSize(FONT_SIZES.title);
    pdf.text(config.title, pageWidth / 2, headerY, { align: 'center' });
    headerY += 8; // Reduced from 10
    if (config.organization) {
        pdf.setFontSize(FONT_SIZES.subtitle);
        pdf.text(config.organization, pageWidth / 2, headerY, { align: 'center' });
        headerY += 6; // space for optional period line
    }
    if (config.period) {
        pdf.setFontSize(FONT_SIZES.small);
        pdf.text(`Report Period: ${config.period}`, pageWidth / 2, headerY, { align: 'center' });
    }

    pdf.setTextColor(0, 0, 0);

    // Start Y position after the header
    let y = headerHeight + 15;

    // Reuse the statistics already gathered

    // Create stats grid similar to web app
    const statSpacing = 5;
    const statHeight = 25;
    const statWidth = (usableWidth - (2 * statSpacing)) / 3;

    // Draw stat cards
    const statData = [
        { label: 'Total Updates', value: reportStats.totalUpdates, color: [0, 120, 212] },
        { label: 'Critical Updates', value: reportStats.criticalUpdates, color: [220, 53, 69] },
        { label: 'Important Updates', value: reportStats.importantUpdates, color: [253, 126, 20] },
        { label: 'Missing Deployments', value: reportStats.totalMissing, color: [108, 117, 125] },
        { label: 'Compliance Rate', value: `${reportStats.complianceRate}%`, color: [40, 167, 69] }
    ];

    const rows = [statData.slice(0, 3), statData.slice(3)];

    rows.forEach((rowStats, rowIndex) => {
        const rowWidth = rowStats.length * statWidth + (rowStats.length - 1) * statSpacing;
        const startX = pageMargin + (usableWidth - rowWidth) / 2;

        rowStats.forEach((stat, index) => {
            const x = startX + (index * (statWidth + statSpacing));

            // Draw card background
            pdf.setFillColor(248, 249, 250);
            pdf.setDrawColor(222, 226, 230);
            pdf.rect(x, y, statWidth, statHeight, 'FD');

            // Draw stat value
            pdf.setFontSize(16);
            pdf.setFont(undefined, 'bold');
            pdf.setTextColor(...stat.color);
            pdf.text(stat.value.toString(), x + statWidth/2, y + 10, { align: 'center' });

            // Draw stat label
            pdf.setFontSize(FONT_SIZES.caption);
            pdf.setFont(undefined, 'normal');
            pdf.setTextColor(108, 117, 125);

            // Split label into lines if needed
            const labelLines = pdf.splitTextToSize(stat.label, statWidth - 4);
            labelLines.forEach((line, lineIndex) => {
                pdf.text(line, x + statWidth/2, y + 16 + (lineIndex * 4), { align: 'center' });
            });
        });

        y += statHeight + 10; // Space between rows
    });

    // Add spacing before the next section
    y += sectionSpacing;

    // Update Age Distribution - keep within margins
    if (y + 35 < pageHeight - 40) {
        pdf.setFontSize(FONT_SIZES.heading);
        pdf.setFont(undefined, 'bold');
        pdf.text('UPDATE AGE ANALYSIS', pageMargin, y);

        const ageColors = {
            '0-7 days': [40, 167, 69],
            '8-14 days': [32, 201, 151],
            '15-30 days': [255, 193, 7],
            '31-60 days': [253, 126, 20],
            '>60 days': [220, 53, 69]
        };

        let ageY = y + 6;
        const ageIndent = 8;
        const ageBarHeight = 3;
        const maxAgeCount = Math.max(...Object.values(updateAges));

        Object.entries(updateAges).forEach(([ageRange, count]) => {
            pdf.setFontSize(FONT_SIZES.caption);
            pdf.setFont(undefined, 'normal');
            const labelX = pageMargin + ageIndent;
            pdf.text(ageRange, labelX, ageY + 2);

            const barStartX = pageMargin + ageIndent + 40;
            const maxBarWidth = usableWidth - ageIndent - 50;
            const barWidth = maxAgeCount > 0 ? (count / maxAgeCount) * maxBarWidth : 0;
            
            pdf.setFillColor(...ageColors[ageRange]);
            pdf.rect(barStartX, ageY, barWidth, ageBarHeight, 'F');
            
            pdf.text(`${count}`, barStartX + barWidth + 5, ageY + 2);
            
            ageY += 5;
        });

    // Add breathing room after the age analysis section
    y = ageY + sectionSpacing;
    }

    // Top Critical Updates - move to first page if space allows
    if (topCriticalUpdates.length > 0 && y + 50 < pageHeight - 30) {
        pdf.setFontSize(FONT_SIZES.heading);
        pdf.setFont(undefined, 'bold');
        pdf.text('TOP CRITICAL UPDATES', pageMargin, y);

        y += 6;

        // Indent the entire list slightly from the left margin
        const bulletIndent = 8;
        const bulletX = pageMargin + bulletIndent;

        const dateX = pageWidth - pageMargin;
        const nameX = bulletX + 4; // space after the number
        const nameWidth = dateX - nameX - 10; // leave ample gap before date column
        topCriticalUpdates.forEach((update, index) => {
            pdf.setFontSize(FONT_SIZES.caption);
            pdf.setFont(undefined, 'bold');
            pdf.text(`${index + 1}.`, bulletX, y);

            pdf.setFont(undefined, 'normal');
            const lines = pdf.splitTextToSize(update.updateName, nameWidth);
            lines.forEach((line, lineIndex) => {
                pdf.text(line, nameX, y + lineIndex * lineHeight);
            });

            const dateY = y + (lines.length - 1) * lineHeight;
            pdf.text(update.releaseDate || '', dateX, dateY, { align: 'right' });

            y += lines.length * lineHeight;
        });

        // Give extra space after the critical updates section
        y += sectionSpacing;
    }

    // Executive summary paragraph (moved below critical updates)
    const summaryData = getExecutiveSummaryData();
    let summaryText = `Update compliance is currently at ${summaryData.complianceRate}% across ${summaryData.totalDevices} released updates, reflecting our ongoing enforcement of patching standards.`;
    if (summaryData.topCritical.length > 0) {
        summaryText += ` This cycle includes several critical updates, such as ${summaryData.topCritical.join(', ')}, all of which are addressed as part of Intune Maintenance Plan.`;
    }
    pdf.setFontSize(FONT_SIZES.body);
    pdf.setFont(undefined, 'normal');
    const summaryLines = pdf.splitTextToSize(summaryText, usableWidth);
    summaryLines.forEach(line => {
        pdf.text(line, pageMargin, y);
        y += lineHeight;
    });

    // Add spacing before the next section
    y += sectionSpacing;

    // Security Severity Distribution chart (dedicated page)
    if (document.getElementById('includeSeverityChart').checked && charts.severity) {
        pdf.addPage();
        currentPage++;
        y = 25; // increased top margin for clearer spacing

        // Page header
        pdf.setFillColor(240, 248, 255);
        pdf.rect(0, 0, pageWidth, 15, 'F');
        pdf.setDrawColor(0, 120, 212);
        pdf.setLineWidth(0.5);
        pdf.line(0, 15, pageWidth, 15);
        pdf.setFontSize(FONT_SIZES.body);
        pdf.setFont(undefined, 'bold');
        pdf.setTextColor(51, 51, 51);
        pdf.text('Charts & Analysis', 15, 10);
        pdf.setFontSize(FONT_SIZES.small);
        pdf.setFont(undefined, 'normal');
        pdf.setTextColor(108, 117, 125);
        pdf.text(`Page ${currentPage}`, pageWidth - pageMargin, 10, { align: 'right' });
        pdf.setTextColor(0, 0, 0);
        pdf.setFontSize(FONT_SIZES.heading);
        pdf.setFont(undefined, 'bold');
        pdf.setTextColor(0, 78, 120);
        pdf.text('Security Severity Distribution – Summary', 20, y);
        pdf.setTextColor(0, 0, 0);
        y += sectionSpacing;

        const chartX = (pageWidth - chartWidth) / 2;
        await addCompactChartToPDF(pdf, 'severityChart', '', chartX, y, chartWidth, chartHeight);
        y += chartHeight + sectionSpacing;

        pdf.setFontSize(FONT_SIZES.body);
        pdf.setFont(undefined, 'normal');
        const lines1 = pdf.splitTextToSize('This chart shows the distribution of updates by security severity for the current period:', usableWidth);
        lines1.forEach(line => {
            pdf.text(line, 20, y);
            y += lineHeight;
        });
        y += paragraphSpacing;

        pdf.setFont(undefined, 'bold');
        const criticalCount = reportStats.criticalUpdates;
        const criticalPercent = Math.round((criticalCount / reportStats.totalUpdates) * 100);
        pdf.setTextColor(220, 53, 69);
        pdf.text(`• Critical: ${criticalCount} updates (${criticalPercent}%)`, 20, y);
        y += lineHeight;

        const importantCount = reportStats.importantUpdates;
        const importantPercent = Math.round((importantCount / reportStats.totalUpdates) * 100);
        pdf.setTextColor(253, 126, 20);
        pdf.text(`• Important: ${importantCount} updates (${importantPercent}%)`, 20, y);
        y += lineHeight;

        const otherCount = reportStats.totalUpdates - criticalCount - importantCount;
        const otherPercent = Math.round((otherCount / reportStats.totalUpdates) * 100);
        pdf.setTextColor(108, 117, 125);
        pdf.text(`• Unspecified or Low: ${otherCount} updates (${otherPercent}%)`, 20, y);
        y += sectionSpacing;

        pdf.setTextColor(0, 0, 0);
        pdf.setFont(undefined, 'normal');
        pdf.setFontSize(FONT_SIZES.small);
        const cvssExplanation = [
            'The severity of each update is assessed using the Common Vulnerability Scoring System (CVSS).',
            'Critical updates are typically reserved for vulnerabilities that could allow remote code execution without user interaction.',
            'Important updates address serious but less severe issues, such as privilege escalation or denial of service.',
            'Unspecified or Low refers to updates that either lack a published CVSS score or are not classified as high-impact. These often include updates related to feature improvements or pending vendor analysis.'
        ];
        cvssExplanation.forEach(text => {
            const lines = pdf.splitTextToSize(text, usableWidth);
            lines.forEach(line => {
                pdf.text(line, 20, y);
                y += lineHeight;
            });
            y += paragraphSpacing;
        });
    }
    // Deployment Summary chart (dedicated page)
    if (document.getElementById('includeDeploymentChart').checked && charts.deployment) {
        pdf.addPage();
        currentPage++;
        y = 25; // increased top margin for clearer spacing

        // Page header
        pdf.setFillColor(240, 248, 255);
        pdf.rect(0, 0, pageWidth, 15, 'F');
        pdf.setDrawColor(0, 120, 212);
        pdf.setLineWidth(0.5);
        pdf.line(0, 15, pageWidth, 15);
        pdf.setFontSize(FONT_SIZES.body);
        pdf.setFont(undefined, 'bold');
        pdf.setTextColor(51, 51, 51);
        pdf.text('Charts & Analysis', 15, 10);
        pdf.setFontSize(FONT_SIZES.small);
        pdf.setFont(undefined, 'normal');
        pdf.setTextColor(108, 117, 125);
        pdf.text(`Page ${currentPage}`, pageWidth - pageMargin, 10, { align: 'right' });
        pdf.setTextColor(0, 0, 0);
        pdf.setFontSize(FONT_SIZES.heading);
        pdf.setFont(undefined, 'bold');
        pdf.setTextColor(0, 78, 120);
        pdf.text('Deployment Summary – Top 10 Updates by Volume', 20, y);
        pdf.setTextColor(0, 0, 0);
        y += sectionSpacing;

        const chartX2 = (pageWidth - chartWidth) / 2;
        await addCompactChartToPDF(pdf, 'deploymentChart', '', chartX2, y, chartWidth, chartHeight);
        y += chartHeight + sectionSpacing;

        pdf.setFontSize(FONT_SIZES.body);
        pdf.setFont(undefined, 'normal');
        const deploymentIntro = pdf.splitTextToSize('This chart shows the deployment status of the top 10 updates with the highest number of assigned devices:', usableWidth);
        deploymentIntro.forEach(line => {
            pdf.text(line, 20, y);
            y += lineHeight;
        });
        y += paragraphSpacing;

        pdf.setFont(undefined, 'bold');
        pdf.text(`• Total missing deployments: ${reportStats.totalMissing}`, 20, y);
        y += lineHeight;
        pdf.text(`• Overall compliance rate: ${reportStats.complianceRate}%`, 20, y);
        y += sectionSpacing;

        pdf.setFont(undefined, 'normal');
        pdf.setFontSize(FONT_SIZES.small);
        const topMissingUpdate = csvData
            .map(row => ({
                name: row['Update Name'],
                missing: parseInt(row['Updates Missing']?.toString().replace(/\D/g, '') || '0')
            }))
            .sort((a, b) => b.missing - a.missing)[0];

        const deploymentExplanation = [
            'This chart highlights updates with the largest number of total targets (both successful and pending). Green bars represent successfully deployed updates, while red bars indicate devices where the update is still pending or failed.',
            topMissingUpdate && topMissingUpdate.missing > 0 ?
                `High-miss updates such as "${topMissingUpdate.name.substring(0, 60)}${topMissingUpdate.name.length > 60 ? '...' : ''}" may require review of deployment assignments, application version conflicts, or device connectivity status.` :
                'All updates show good deployment coverage with minimal missing installations.',
            'This data is used to help identify:',
            '• Trends in common deployment failures',
            '• Apps with wide install footprints but low coverage',
            '• Potential priorities for remediation in the next patch cycle'
        ];
        deploymentExplanation.forEach(text => {
            const lines = pdf.splitTextToSize(text, usableWidth);
            lines.forEach(line => {
                pdf.text(line, text.startsWith('•') ? 25 : 20, y);
                y += lineHeight;
            });
            y += paragraphSpacing;
        });
    }



    // Check if we need a new page for trend chart
    if (y > pageHeight - 100) {
        pdf.addPage();
        currentPage++;
        y = 25; // increased top margin for clearer spacing
        
        // Page header
        pdf.setFillColor(240, 248, 255);
        pdf.rect(0, 0, pageWidth, 15, 'F');
        pdf.setDrawColor(0, 120, 212);
        pdf.setLineWidth(0.5);
        pdf.line(0, 15, pageWidth, 15);
        pdf.setFontSize(FONT_SIZES.body);
        pdf.setFont(undefined, 'bold');
        pdf.setTextColor(51, 51, 51);
        pdf.text('Charts & Analysis (continued)', 15, 10);
        pdf.setFontSize(FONT_SIZES.small);
        pdf.setFont(undefined, 'normal');
        pdf.setTextColor(108, 117, 125);
        pdf.text(`Page ${currentPage}`, pageWidth - pageMargin, 10, { align: 'right' });
        pdf.setTextColor(0, 0, 0);
    }
    
    // Trend Chart
    if (document.getElementById('includeTrendChart').checked && charts.trend) {
        // Chart title
        pdf.setFontSize(FONT_SIZES.heading);
        pdf.setFont(undefined, 'bold');
        pdf.setTextColor(0, 78, 120);
        pdf.text('Update Release Timeline – Overview', 20, y);
        pdf.setTextColor(0, 0, 0);
        y += sectionSpacing;
        
        const chartX = (pageWidth - chartWidth) / 2;
        await addCompactChartToPDF(pdf, 'trendChart', '', chartX, y, chartWidth, chartHeight);
        y += chartHeight + sectionSpacing;
        
        // Explanation
        pdf.setFontSize(FONT_SIZES.body);
        pdf.setFont(undefined, 'normal');
        
        const monthCounts = {};
        csvData.forEach(row => {
            const dateStr = row['Release Date'];
            if (dateStr) {
                try {
                    const date = new Date(dateStr.replace(/-/g, '/'));
                    if (!isNaN(date.getTime())) {
                        const monthKey = date.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
                        monthCounts[monthKey] = (monthCounts[monthKey] || 0) + 1;
                    }
                } catch (e) {
                    // Ignore parse errors
                }
            }
        });
        
        const sortedMonths = Object.keys(monthCounts).sort((a, b) => new Date(a) - new Date(b));
        const latestMonth = sortedMonths[sortedMonths.length - 1];
        const latestCount = monthCounts[latestMonth] || 0;
        const firstMonth = sortedMonths[0];
        
        let trendDescription = 'Consistent month-over-month';
        if (sortedMonths.length >= 3) {
            const recentCounts = sortedMonths.slice(-3).map(m => monthCounts[m]);
            const avgRecent = recentCounts.reduce((a, b) => a + b, 0) / recentCounts.length;
            const olderCounts = sortedMonths.slice(0, -3).map(m => monthCounts[m]);
            const avgOlder = olderCounts.length > 0 ? olderCounts.reduce((a, b) => a + b, 0) / olderCounts.length : avgRecent;
            
            if (avgRecent > avgOlder * 1.3) {
                trendDescription = 'Significant increase';
            } else if (avgRecent > avgOlder * 1.1) {
                trendDescription = 'Gradual rise';
            } else if (avgRecent < avgOlder * 0.7) {
                trendDescription = 'Notable decrease';
            }
        }
        
        const trendIntro = pdf.splitTextToSize('This graph displays the trend of updates released over time, highlighting our ongoing maintenance and security efforts:', usableWidth);
        trendIntro.forEach(line => {
            pdf.text(line, 20, y);
            y += lineHeight;
        });
        y += paragraphSpacing;
        
        // Trend stats
        pdf.setFont(undefined, 'bold');
        pdf.text(`• Update count (latest month): ${latestCount}`, 20, y);
        y += lineHeight;
        pdf.text(`• Trend: ${trendDescription}`, 20, y);
        y += lineHeight;
        pdf.text(`• Observation period: ${firstMonth} to ${latestMonth}`, 20, y);
        y += sectionSpacing;
        
        pdf.setFont(undefined, 'normal');
        pdf.setFontSize(FONT_SIZES.small);
        
        const trendExplanation = [
            'All updates are reviewed and addressed as part of our structured patch management process.',
            '• Releases follow Microsoft\'s Patch Tuesday schedule and include both first-party and third-party updates.',
            '• Elevated activity in certain months corresponds to major cumulative or security update cycles.',
            '• A consistent release pattern confirms that systems are being actively maintained in alignment with organizational policies and security standards.',
            '• Some older updates may show up because new devices were scanned, systems were reset, or additional updates were identified during inventory checks.',
            'We continue to take timely action on all identified updates to maintain compliance and minimize risk across your environment as part of your Intune Maintenance plan.'
        ];
        
        trendExplanation.forEach((text, index) => {
            if (y > pageHeight - 25) {
                pdf.addPage();
                currentPage++;
                y = 25; // increased top margin for clearer spacing
                
                // Page header
                pdf.setFillColor(240, 248, 255);
                pdf.rect(0, 0, pageWidth, 15, 'F');
                pdf.setDrawColor(0, 120, 212);
                pdf.setLineWidth(0.5);
                pdf.line(0, 15, pageWidth, 15);
                pdf.setFontSize(FONT_SIZES.body);
                pdf.setFont(undefined, 'bold');
                pdf.setTextColor(51, 51, 51);
                pdf.text('Charts & Analysis (continued)', 15, 10);
                pdf.setFontSize(FONT_SIZES.small);
                pdf.setFont(undefined, 'normal');
                pdf.setTextColor(108, 117, 125);
                pdf.text(`Page ${currentPage}`, pageWidth - pageMargin, 10, { align: 'right' });
                pdf.setTextColor(0, 0, 0);
                pdf.setFontSize(FONT_SIZES.small);
                pdf.setFont(undefined, 'normal');
            }
            
            if (index === 0) {
                pdf.setFont(undefined, 'bold');
            } else {
                pdf.setFont(undefined, 'normal');
            }
            
            const lines = pdf.splitTextToSize(text, usableWidth);
            lines.forEach(line => {
                pdf.text(line, text.startsWith('•') ? 25 : 20, y);
                y += lineHeight;
            });
            y += paragraphSpacing;
        });
        y += sectionSpacing;
    }

    // Detailed table
    if (document.getElementById('includeDetailTable').checked) {
        updateProgress(70, 'Adding detailed table...');
        addDetailedTable(pdf, FONT_SIZES);
    }
    
    updateProgress(90, 'Finalizing report...');
}

async function addCompactChartToPDF(pdf, chartId, title, x, y, width, height) {
    // Title is now added separately, so skip if empty
    if (title) {
        pdf.setFontSize(10);
        pdf.setFont(undefined, 'bold');
        pdf.text(title, x, y - 2);
    }
    
    // Get chart canvas
    const canvas = document.getElementById(chartId);
    if (!canvas) return;
    
    // Wait for chart to render
    await new Promise(resolve => setTimeout(resolve, 100));
    
    // Convert chart to image
    const imgData = canvas.toDataURL('image/png');
    
    // Calculate proper dimensions maintaining aspect ratio
    const aspectRatio = canvas.width / canvas.height;
    let imgWidth = width;
    let imgHeight = height;
    
    if (imgWidth / imgHeight > aspectRatio) {
        imgWidth = imgHeight * aspectRatio;
    } else {
        imgHeight = imgWidth / aspectRatio;
    }
    
    // Add image to PDF
    pdf.addImage(imgData, 'PNG', x, y, imgWidth, imgHeight);
}

function addDetailedTable(pdf, FONT_SIZES) {
    pdf.addPage();
    const tableStartPage = pdf.internal.getNumberOfPages();
    
    // Enhanced header with gradient effect
    pdf.setFillColor(240, 248, 255);
    pdf.rect(0, 0, pdf.internal.pageSize.getWidth(), 15, 'F');
    
    // Add accent line
    pdf.setDrawColor(0, 120, 212);
    pdf.setLineWidth(0.5);
    pdf.line(0, 15, pdf.internal.pageSize.getWidth(), 15);
    
    pdf.setFontSize(FONT_SIZES.body);
    pdf.setFont(undefined, 'bold');
    pdf.setTextColor(51, 51, 51);
    pdf.text('Detailed Update List', 15, 10);
    
    // Add page number
    pdf.setFontSize(FONT_SIZES.small);
    pdf.setFont(undefined, 'normal');
    pdf.setTextColor(108, 117, 125);
    pdf.text(`Page ${tableStartPage}`, pdf.internal.pageSize.getWidth() - 20, 10, { align: 'right' });
    
    pdf.setTextColor(0, 0, 0);
    
    // Add table description
    pdf.setFontSize(FONT_SIZES.small);
    pdf.setFont(undefined, 'normal');
    pdf.text('Complete listing of all Windows updates tracked during this reporting period.', 15, 22);
    pdf.text('Color coding: Critical (Red), Important (Orange), Moderate (Yellow), Low (Green)', 15, 26);

    const sortedData = getSortedData();
    const highMissingRows = [];
    const oldRows = [];

    let tableData = sortedData.map((row, idx) => {
        const missing = parseInt(row['Updates Missing']?.toString().replace(/\D/g, '') || '0');
        const daysOld = calculateDaysOld(row['Release Date']);
        if (missing > 10) highMissingRows.push(idx);
        if (daysOld > 30) oldRows.push(idx);

        return [
            row['Update Name'] || '',
            row['Version'] || '',
            row['Security Severity'] || 'Unspecified',
            row['Release Date'] || '',
            row['Updates Missing'] || '0',
            row['Recently Deployed'] || '0',
            getStatusWithReboot(row)
        ];
    });
    
    // Add table with corrected column widths
    if (typeof pdf.autoTable === 'function') {
        const sideMargin = 15;
        pdf.autoTable({
            head: [['Update Name', 'Version', 'Severity', 'Date', 'Pending Updates', 'Deployed', 'Status']],
            body: tableData,
            startY: 32,
            theme: 'striped',
            headStyles: {
                fillColor: [0, 120, 212],
                fontSize: FONT_SIZES.caption,
                fontStyle: 'bold',
                cellPadding: 1.5,
                halign: 'center'
            },
            bodyStyles: {
                fontSize: FONT_SIZES.caption,
                cellPadding: 1.5
            },
            alternateRowStyles: {
                fillColor: [248, 249, 250]
            },
            // Dynamically calculate column widths so the table spans the
            // printable area without exceeding the left/right margins.
            columnStyles: (function () {
                const pageWidth = pdf.internal.pageSize.getWidth();
                const tableWidth = pageWidth - sideMargin * 2;
                const ratios = [0.333, 0.113, 0.113, 0.113, 0.082, 0.082, 0.164];
                return ratios.reduce((styles, ratio, idx) => {
                    styles[idx] = {
                        cellWidth: tableWidth * ratio,
                        halign: idx === 0 ? 'left' : 'center'
                    };
                    if (idx === 0 || idx === 6) styles[idx].overflow = 'linebreak';
                    if (idx === 4 || idx === 5) styles[idx].fontStyle = 'bold';
                    return styles;
                }, {});
            })(),
            margin: { top: 20, left: sideMargin, right: sideMargin },
            showHead: 'everyPage',
            tableWidth: pdf.internal.pageSize.getWidth() - sideMargin * 2,
            didDrawPage: function(data) {
                // Add page header on continuation pages
                if (data.pageNumber > 1) {
                    pdf.setFillColor(240, 248, 255);
                    pdf.rect(0, 0, pdf.internal.pageSize.getWidth(), 15, 'F');
                
                pdf.setDrawColor(0, 120, 212);
                pdf.setLineWidth(0.5);
                pdf.line(0, 15, pdf.internal.pageSize.getWidth(), 15);
                
                pdf.setFontSize(FONT_SIZES.body);
                pdf.setFont(undefined, 'bold');
                pdf.setTextColor(51, 51, 51);
                pdf.text('Detailed Update List (continued)', 15, 10);
                
                const currentPage = tableStartPage + data.pageNumber - 1;
                pdf.setFontSize(FONT_SIZES.small);
                pdf.setFont(undefined, 'normal');
                pdf.setTextColor(108, 117, 125);
                pdf.text(`Page ${currentPage}`, pdf.internal.pageSize.getWidth() - 20, 10, { align: 'right' });
                
                pdf.setTextColor(0, 0, 0);
            }
        },
        didParseCell: function(data) {
            if (data.row.section === 'body') {
                if (highMissingRows.includes(data.row.index)) {
                    data.cell.styles.fillColor = [248, 215, 218];
                } else if (oldRows.includes(data.row.index)) {
                    data.cell.styles.fillColor = [255, 243, 205];
                }
            }
            // Color code severity cells
            if (data.column.index === 2 && data.cell.raw) {
                const severity = data.cell.raw.toLowerCase();
                if (severity.includes('critical')) {
                    data.cell.styles.textColor = [220, 53, 69];
                    data.cell.styles.fontStyle = 'bold';
                } else if (severity.includes('important')) {
                    data.cell.styles.textColor = [253, 126, 20];
                    data.cell.styles.fontStyle = 'bold';
                } else if (severity.includes('moderate')) {
                    data.cell.styles.textColor = [255, 193, 7];
                } else if (severity.includes('low')) {
                    data.cell.styles.textColor = [40, 167, 69];
                }
            }
            
            // Highlight high missing counts
            if (data.column.index === 4 && data.cell.raw) {
                const missing = parseInt(data.cell.raw);
                if (missing > 20) {
                    data.cell.styles.textColor = [220, 53, 69];
                } else if (missing > 10) {
                    data.cell.styles.textColor = [253, 126, 20];
                }
            }
            
            // Highlight high deployment counts
            if (data.column.index === 5 && data.cell.raw) {
                const deployed = parseInt(data.cell.raw);
                if (deployed >= 90) {
                    data.cell.styles.textColor = [40, 167, 69];
                }
            }
        }
    });
    } else {
        console.warn('jsPDF AutoTable plugin not loaded, skipping details table');
    }
    
    // Add summary at the end of the table
    const finalY = pdf.lastAutoTable ? (pdf.lastAutoTable.finalY || pdf.internal.pageSize.getHeight() - 50) : pdf.internal.pageSize.getHeight() - 50;
    if (finalY < pdf.internal.pageSize.getHeight() - 30) {
        pdf.setFillColor(248, 249, 250);
        pdf.rect(15, finalY + 10, pdf.internal.pageSize.getWidth() - 30, 18, 'F');

        pdf.setFontSize(FONT_SIZES.caption);
        pdf.setFont(undefined, 'italic');
        pdf.text(`Table contains ${csvData.length} total updates sorted by severity and pending count.`, 18, finalY + 18);
        pdf.text('Rows in red exceed 10 pending devices; yellow rows are older than 30 days.', 18, finalY + 25);
    }
}

function updateProgress(percentage, text) {
    const progressFill = document.getElementById('progressFill');
    const statusText = document.getElementById('statusText');
    
    progressFill.style.width = percentage + '%';
    progressFill.textContent = percentage + '%';
    statusText.textContent = text;
}

function downloadSampleCSV() {
    const sampleData = [
        ['Update Name', 'Version', 'Security Severity', 'Release Date', 'Updates Missing', 'Recently Deployed', 'Details'],
        ['Windows Security Update KB5034441', '2025-01', 'Critical', '2025-01-10', '15', '85', '1'],
        ['Microsoft Defender Definition Update', '1.403.2956.0', 'Important', '2025-01-12', '3', '97', '1'],
        ['Microsoft Edge Security Update', '120.0.2210.144', 'Important', '2025-01-08', '8', '92', '1'],
        ['Windows Cumulative Update KB5034439', '2025-01', 'Moderate', '2025-01-10', '22', '78', '1'],
        ['.NET Framework Security Update', '4.8.1', 'Important', '2025-02-09', '12', '88', '1'],
        ['Office Security Update KB5035552', '16.0.17425', 'Critical', '2025-02-11', '5', '95', '1'],
        ['Windows Defender Update', '1.404.123.0', 'Important', '2025-02-14', '10', '90', '1'],
        ['Visual C++ Redistributable Update', '14.38.33130', 'Moderate', '2024-12-05', '30', '70', '1'],
        ['Windows Malicious Software Removal Tool', '5.119', 'Low', '2024-12-12', '0', '100', '1'],
        ['SQL Server Security Update', '15.0.4360.1', 'Critical', '2024-12-12', '18', '82', '1']
    ];
    
    const csvContent = sampleData.map(row => row.join(',')).join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = 'sample_update_data.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function showAlert(message, type) {
    const alertDiv = document.createElement('div');
    alertDiv.className = `alert alert-${type}`;
    alertDiv.textContent = message;
    
    const mainContent = document.querySelector('.main-content');
    mainContent.insertBefore(alertDiv, mainContent.firstChild);
    
    setTimeout(() => {
        alertDiv.remove();
    }, 5000);
}

function saveFormData() {
    const formData = {
        reportTitle: document.getElementById('reportTitle').value,
        organizationName: document.getElementById('organizationName').value,
        reportPeriod: document.getElementById('reportPeriod').value,
        reportTheme: document.getElementById('reportTheme').value,
        additionalNotes: document.getElementById('additionalNotes').value,
        options: {
            includeSummary: document.getElementById('includeSummary').checked,
            includeSeverityChart: document.getElementById('includeSeverityChart').checked,
            includeDeploymentChart: document.getElementById('includeDeploymentChart').checked,
            includeTrendChart: document.getElementById('includeTrendChart').checked,
            includeDetailTable: document.getElementById('includeDetailTable').checked,
            includeRecommendations: document.getElementById('includeRecommendations').checked
        }
    };
    localStorage.setItem('updateReportFormData', JSON.stringify(formData));
}

function loadSavedPreferences() {
    // Auto-populate current month
    const currentMonth = new Date().toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    document.getElementById('reportPeriod').value = currentMonth;
    
    // Load saved form data
    const saved = localStorage.getItem('updateReportFormData');
    if (saved) {
        try {
            const formData = JSON.parse(saved);
            if (formData.organizationName) {
                document.getElementById('organizationName').value = formData.organizationName;
            }
        } catch (e) {
            console.error('Error loading saved form data:', e);
        }
    }
}

console.log('Windows Update Report Generator initialized successfully');
