/**
 * Excel Lite Chart Manager
 * Provides chart creation and management functionality
 */
class ChartManager {
    constructor(spreadsheet) {
        this.spreadsheet = spreadsheet;
        this.charts = [];
        this.chartCounter = 0;
        this.chartColors = [
            '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', 
            '#5B9BD5', '#70AD47', '#7030A0', '#C00000'
        ];
    }
    
    /**
     * Create a new chart
     * @param {string} type - Chart type (bar, column, line, pie)
     * @param {string} dataRange - Cell range (e.g., "A1:B5")
     * @param {object} options - Chart options
     * @returns {number} - Chart ID
     */
    createChart(type, dataRange, options = {}) {
        // Parse range to get data
        const [startCell, endCell] = dataRange.split(':').map(cell => cell.trim());
        const data = this.extractDataFromRange(startCell, endCell);
        
        if (!data || data.values.length === 0) {
            console.error('Invalid data range for chart');
            return null;
        }
        
        const chartId = ++this.chartCounter;
        
        const chart = {
            id: chartId,
            type: type,
            title: options.title || `Chart ${chartId}`,
            range: dataRange,
            data: data,
            options: {
                showLegend: options.showLegend !== undefined ? options.showLegend : true,
                showLabels: options.showLabels !== undefined ? options.showLabels : true,
                colors: options.colors || this.chartColors,
                width: options.width || 400,
                height: options.height || 300
            }
        };
        
        this.charts.push(chart);
        return chartId;
    }
    
    /**
     * Get chart by ID
     * @param {number} chartId - Chart ID
     * @returns {object} - Chart object
     */
    getChart(chartId) {
        return this.charts.find(chart => chart.id === chartId);
    }
    
    /**
     * Update an existing chart
     * @param {number} chartId - Chart ID
     * @param {object} updates - Updates to apply
     */
    updateChart(chartId, updates) {
        const chart = this.getChart(chartId);
        if (!chart) return;
        
        if (updates.type) chart.type = updates.type;
        if (updates.title) chart.title = updates.title;
        
        if (updates.range) {
            chart.range = updates.range;
            const [startCell, endCell] = updates.range.split(':').map(cell => cell.trim());
            chart.data = this.extractDataFromRange(startCell, endCell);
        }
        
        if (updates.options) {
            chart.options = { ...chart.options, ...updates.options };
        }
    }
    
    /**
     * Delete a chart
     * @param {number} chartId - Chart ID
     */
    deleteChart(chartId) {
        const index = this.charts.findIndex(chart => chart.id === chartId);
        if (index !== -1) {
            this.charts.splice(index, 1);
        }
    }
    
    /**
     * Extract data from a cell range
     * @param {string} startCell - Start cell reference (e.g., "A1")
     * @param {string} endCell - End cell reference (e.g., "B5")
     * @returns {object} - Extracted data
     */
    extractDataFromRange(startCell, endCell) {
        const startCoords = this.parseCellReference(startCell);
        const endCoords = this.parseCellReference(endCell);
        
        if (!startCoords || !endCoords) return null;
        
        const values = [];
        const labels = [];
        const series = [];
        
        const sheet = this.spreadsheet.getCurrentSheet();
        
        // Determine if we should use first row as labels
        const hasLabels = true; // We could make this configurable
        
        // Extract data
        for (let row = startCoords.row; row <= endCoords.row; row++) {
            const rowData = [];
            
            for (let col = startCoords.col; col <= endCoords.col; col++) {
                if (sheet.data[row] && sheet.data[row][col] !== undefined) {
                    let value = sheet.data[row][col];
                    
                    // If cell contains a formula, use the display value
                    if (typeof value === 'string' && value.startsWith('=') && 
                        sheet.displayData && sheet.displayData[row] && 
                        sheet.displayData[row][col] !== undefined) {
                        value = sheet.displayData[row][col];
                    }
                    
                    // Convert numeric strings to numbers
                    if (typeof value === 'string' && !isNaN(parseFloat(value))) {
                        value = parseFloat(value);
                    }
                    
                    rowData.push(value);
                } else {
                    rowData.push('');
                }
            }
            
            values.push(rowData);
        }
        
        // Extract labels from first row if hasLabels is true
        if (hasLabels && values.length > 0) {
            labels.push(...values[0].slice(1)); // Assuming first column is categories
        }
        
        // Extract series names from first column
        if (values.length > (hasLabels ? 1 : 0)) {
            for (let i = hasLabels ? 1 : 0; i < values.length; i++) {
                if (values[i].length > 0) {
                    series.push({
                        name: values[i][0].toString(),
                        data: values[i].slice(1)
                    });
                }
            }
        }
        
        // Extract categories from first column if there are no labels
        const categories = [];
        if (values.length > 0) {
            for (let i = hasLabels ? 1 : 0; i < values.length; i++) {
                if (values[i].length > 0) {
                    categories.push(values[i][0]);
                }
            }
        }
        
        return {
            values,
            labels,
            series,
            categories,
            hasLabels
        };
    }
    
    /**
     * Render a chart to a container
     * @param {number} chartId - Chart ID
     * @param {HTMLElement} container - Container element
     */
    renderChart(chartId, container) {
        const chart = this.getChart(chartId);
        if (!chart || !container) return;
        
        // Clear container
        container.innerHTML = '';
        container.style.width = `${chart.options.width}px`;
        container.style.height = `${chart.options.height}px`;
        
        // Create chart title
        const titleElement = document.createElement('div');
        titleElement.className = 'chart-title';
        titleElement.textContent = chart.title;
        container.appendChild(titleElement);
        
        // Create chart canvas
        const canvas = document.createElement('canvas');
        canvas.width = chart.options.width;
        canvas.height = chart.options.height - 40; // Leave space for title
        container.appendChild(canvas);
        
        // Render appropriate chart type
        switch (chart.type.toLowerCase()) {
            case 'bar':
                this.renderBarChart(chart, canvas);
                break;
            case 'column':
                this.renderColumnChart(chart, canvas);
                break;
            case 'line':
                this.renderLineChart(chart, canvas);
                break;
            case 'pie':
                this.renderPieChart(chart, canvas);
                break;
            default:
                console.error(`Unsupported chart type: ${chart.type}`);
        }
        
        // Add legend if needed
        if (chart.options.showLegend) {
            const legendElement = document.createElement('div');
            legendElement.className = 'chart-legend';
            
            chart.data.series.forEach((series, index) => {
                const legendItem = document.createElement('div');
                legendItem.className = 'legend-item';
                
                const colorSwatch = document.createElement('span');
                colorSwatch.className = 'color-swatch';
                colorSwatch.style.backgroundColor = chart.options.colors[index % chart.options.colors.length];
                
                const legendText = document.createElement('span');
                legendText.textContent = series.name;
                
                legendItem.appendChild(colorSwatch);
                legendItem.appendChild(legendText);
                legendElement.appendChild(legendItem);
            });
            
            container.appendChild(legendElement);
        }
    }
    
    /**
     * Simple bar chart renderer
     * In a real implementation, you might use a library like Chart.js
     */
    renderBarChart(chart, canvas) {
        const ctx = canvas.getContext('2d');
        const width = canvas.width;
        const height = canvas.height;
        
        // Clear canvas
        ctx.clearRect(0, 0, width, height);
        
        // Set some margins
        const margin = { top: 20, right: 20, bottom: 30, left: 40 };
        const chartWidth = width - margin.left - margin.right;
        const chartHeight = height - margin.top - margin.bottom;
        
        // Calculate max value for scaling
        let maxValue = 0;
        chart.data.series.forEach(series => {
            series.data.forEach(value => {
                if (!isNaN(value) && value > maxValue) {
                    maxValue = value;
                }
            });
        });
        
        // Add 10% to max value for better visualization
        maxValue *= 1.1;
        
        // Draw axes
        ctx.strokeStyle = '#333';
        ctx.lineWidth = 1;
        
        // Y-axis
        ctx.beginPath();
        ctx.moveTo(margin.left, margin.top);
        ctx.lineTo(margin.left, height - margin.bottom);
        ctx.stroke();
        
        // X-axis
        ctx.beginPath();
        ctx.moveTo(margin.left, height - margin.bottom);
        ctx.lineTo(width - margin.right, height - margin.bottom);
        ctx.stroke();
        
        // Draw bars
        const categoryCount = chart.data.categories.length;
        const seriesCount = chart.data.series.length;
        const barWidth = chartWidth / (categoryCount * seriesCount + categoryCount + 1);
        const groupWidth = barWidth * seriesCount;
        
        chart.data.series.forEach((series, seriesIndex) => {
            ctx.fillStyle = chart.options.colors[seriesIndex % chart.options.colors.length];
            
            series.data.forEach((value, dataIndex) => {
                if (!isNaN(value)) {
                    const barHeight = (value / maxValue) * chartHeight;
                    const x = margin.left + (dataIndex + 1) * barWidth + dataIndex * groupWidth + seriesIndex * barWidth;
                    const y = height - margin.bottom - barHeight;
                    
                    ctx.fillRect(x, y, barWidth * 0.8, barHeight);
                    
                    // Add label if enabled
                    if (chart.options.showLabels) {
                        ctx.fillStyle = '#333';
                        ctx.font = '10px Arial';
                        ctx.textAlign = 'center';
                        ctx.fillText(value.toString(), x + barWidth * 0.4, y - 5);
                        ctx.fillStyle = chart.options.colors[seriesIndex % chart.options.colors.length];
                    }
                }
            });
        });
        
        // Draw category labels
        ctx.fillStyle = '#333';
        ctx.font = '10px Arial';
        ctx.textAlign = 'center';
        
        chart.data.categories.forEach((category, index) => {
            const x = margin.left + (index + 1) * barWidth + index * groupWidth + groupWidth / 2;
            ctx.fillText(category.toString(), x, height - margin.bottom + 15);
        });
    }
    
    renderColumnChart(chart, canvas) {
        // Similar to bar chart but with horizontal bars
        // Implementation would be similar to renderBarChart but with x and y swapped
        this.renderBarChart(chart, canvas); // For simplicity, reuse bar chart
    }
    
    renderLineChart(chart, canvas) {
        const ctx = canvas.getContext('2d');
        const width = canvas.width;
        const height = canvas.height;
        
        // Clear canvas
        ctx.clearRect(0, 0, width, height);
        
        // Set some margins
        const margin = { top: 20, right: 20, bottom: 30, left: 40 };
        const chartWidth = width - margin.left - margin.right;
        const chartHeight = height - margin.top - margin.bottom;
        
        // Calculate max value for scaling
        let maxValue = 0;
        chart.data.series.forEach(series => {
            series.data.forEach(value => {
                if (!isNaN(value) && value > maxValue) {
                    maxValue = value;
                }
            });
        });
        
        // Add 10% to max value for better visualization
        maxValue *= 1.1;
        
        // Draw axes
        ctx.strokeStyle = '#333';
        ctx.lineWidth = 1;
        
        // Y-axis
        ctx.beginPath();
        ctx.moveTo(margin.left, margin.top);
        ctx.lineTo(margin.left, height - margin.bottom);
        ctx.stroke();
        
        // X-axis
        ctx.beginPath();
        ctx.moveTo(margin.left, height - margin.bottom);
        ctx.lineTo(width - margin.right, height - margin.bottom);
        ctx.stroke();
        
        // Draw lines
        const categoryCount = chart.data.categories.length;
        const pointWidth = chartWidth / (categoryCount + 1);
        
        chart.data.series.forEach((series, seriesIndex) => {
            ctx.strokeStyle = chart.options.colors[seriesIndex % chart.options.colors.length];
            ctx.lineWidth = 2;
            
            ctx.beginPath();
            
            series.data.forEach((value, dataIndex) => {
                if (!isNaN(value)) {
                    const x = margin.left + (dataIndex + 1) * pointWidth;
                    const y = height - margin.bottom - (value / maxValue) * chartHeight;
                    
                    if (dataIndex === 0) {
                        ctx.moveTo(x, y);
                    } else {
                        ctx.lineTo(x, y);
                    }
                    
                    // Draw points
                    ctx.fillStyle = chart.options.colors[seriesIndex % chart.options.colors.length];
                    ctx.beginPath();
                    ctx.arc(x, y, 4, 0, Math.PI * 2);
                    ctx.fill();
                    
                    // Add label if enabled
                    if (chart.options.showLabels) {
                        ctx.fillStyle = '#333';
                        ctx.font = '10px Arial';
                        ctx.textAlign = 'center';
                        ctx.fillText(value.toString(), x, y - 10);
                    }
                }
            });
            
            ctx.stroke();
        });
        
        // Draw category labels
        ctx.fillStyle = '#333';
        ctx.font = '10px Arial';
        ctx.textAlign = 'center';
        
        chart.data.categories.forEach((category, index) => {
            const x = margin.left + (index + 1) * pointWidth;
            ctx.fillText(category.toString(), x, height - margin.bottom + 15);
        });
    }
    
    renderPieChart(chart, canvas) {
        const ctx = canvas.getContext('2d');
        const width = canvas.width;
        const height = canvas.height;
        
        // Clear canvas
        ctx.clearRect(0, 0, width, height);
        
        // Get the first series for pie chart
        if (chart.data.series.length === 0 || !chart.data.series[0].data) {
            return;
        }
        
        const data = chart.data.series[0].data;
        const labels = chart.data.labels;
        
        // Calculate total
        let total = 0;
        data.forEach(value => {
            if (!isNaN(value)) {
                total += value;
            }
        });
        
        // Draw pie
        const radius = Math.min(width, height) * 0.4;
        const centerX = width / 2;
        const centerY = height / 2;
        
        let startAngle = 0;
        
        data.forEach((value, index) => {
            if (!isNaN(value) && value > 0) {
                const sliceAngle = (value / total) * Math.PI * 2;
                
                ctx.fillStyle = chart.options.colors[index % chart.options.colors.length];
                
                ctx.beginPath();
                ctx.moveTo(centerX, centerY);
                ctx.arc(centerX, centerY, radius, startAngle, startAngle + sliceAngle);
                ctx.closePath();
                ctx.fill();
                
                // Add label if enabled
                if (chart.options.showLabels) {
                    // Position label in the middle of the slice
                    const labelAngle = startAngle + sliceAngle / 2;
                    const labelRadius = radius * 0.7;
                    const labelX = centerX + Math.cos(labelAngle) * labelRadius;
                    const labelY = centerY + Math.sin(labelAngle) * labelRadius;
                    
                    ctx.fillStyle = '#fff';
                    ctx.font = 'bold 12px Arial';
                    ctx.textAlign = 'center';
                    ctx.textBaseline = 'middle';
                    
                    const percentage = Math.round((value / total) * 100);
                    ctx.fillText(`${percentage}%`, labelX, labelY);
                }
                
                startAngle += sliceAngle;
            }
        });
    }
    
    /**
     * Parse cell reference to row and column indices
     * @param {string} cellRef - Cell reference (e.g., "A1")
     * @returns {object|null} - Row and column indices or null if invalid
     */
    parseCellReference(cellRef) {
        const match = cellRef.match(/^([A-Z]+)([0-9]+)$/);
        if (!match) return null;
        
        const colStr = match[1];
        const rowNum = parseInt(match[2], 10);
        
        // Convert column letters to index (A=0, B=1, ..., Z=25, AA=26, ...)
        let colIndex = 0;
        for (let i = 0; i < colStr.length; i++) {
            colIndex = colIndex * 26 + (colStr.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
        }
        colIndex--; // Adjust to 0-based index
        
        // Convert row number to index (1-based to 0-based)
        const rowIndex = rowNum - 1;
        
        return { row: rowIndex, col: colIndex };
    }
}

// Add Chart Manager to the window object
window.ChartManager = ChartManager;
