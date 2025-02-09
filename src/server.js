const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const fetch = require('node-fetch');
const cors = require('cors');
const path = require('path');
require('dotenv').config();

// Initialize express app
const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Constants for processing
const BATCH_SIZE = 5; // Smaller batch size for better stability
const DELAY_BETWEEN_BATCHES = 1000; // 1 second between batches
const DELAY_BETWEEN_KEYWORDS = 200; // 200ms between keywords
const MAX_RETRIES = 3; // Maximum retries for failed API calls

// SSE Constants
const SSE_KEEPALIVE_INTERVAL = 5000; // 5 seconds
const SSE_TIMEOUT = 2 * 60 * 60 * 1000; // 2 hours
const MAX_ERRORS_BEFORE_PAUSE = 3;
const ERROR_PAUSE_DURATION = 30000; // 30 seconds

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(express.static(path.join(__dirname, '../public')));

// Configure request timeouts
app.use((req, res, next) => {
    res.setTimeout(300000, () => {
        res.status(408).send('Request timeout');
    });
    next();
});

// Store active SSE connections and processing state
const clients = new Map();
let isProcessing = false;
let analysisResults = [];
let currentProcessingState = null;

// Add API key validation
if (!process.env.OPENROUTER_API_KEY) {
    console.error('OPENROUTER_API_KEY is not set in environment variables');
}

// Modified SSE endpoint
app.get('/api/analysis-progress', (req, res) => {
    const clientId = Date.now().toString(36) + Math.random().toString(36).substr(2);
    
    // Configure headers for stable SSE connection
    res.writeHead(200, {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache, no-transform',
        'Connection': 'keep-alive',
        'X-Accel-Buffering': 'no',
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Credentials': 'true'
    });

    // Send initial state if reconnecting during processing
    if (currentProcessingState) {
        res.write(`data: ${JSON.stringify(currentProcessingState)}\n\n`);
    }

    // Keep-alive interval
    const keepalive = setInterval(() => {
        try {
            res.write(`data: ${JSON.stringify({ 
                type: 'keepalive',
                timestamp: Date.now(),
                isProcessing
            })}\n\n`);
        } catch (error) {
            cleanup();
        }
    }, SSE_KEEPALIVE_INTERVAL);

    // Cleanup function
    const cleanup = () => {
        clearInterval(keepalive);
        clients.delete(clientId);
    };

    // Store client
    clients.set(clientId, {
        res,
        startTime: Date.now(),
        cleanup
    });

    // Auto-cleanup after timeout
    setTimeout(cleanup, SSE_TIMEOUT);

    // Handle disconnection
    req.on('close', cleanup);
});

// Improved sendProgressUpdate function
function sendProgressUpdate(data) {
    // Store current state for reconnecting clients
    currentProcessingState = data;
    
    const message = `data: ${JSON.stringify({
        ...data,
        timestamp: Date.now(),
        isProcessing
    })}\n\n`;

    clients.forEach((client, clientId) => {
        try {
            client.res.write(message);
        } catch (error) {
            client.cleanup();
        }
    });
}

// Improved keyword analysis with retries
async function analyzeKeyword(keyword, matchType, topic, retryCount = 0) {
    try {
        sendProgressUpdate({
            type: 'processing',
            keyword,
            status: 'Processing'
        });

        const response = await fetch('https://openrouter.ai/api/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${process.env.OPENROUTER_API_KEY}`,
                'Content-Type': 'application/json',
                'HTTP-Referer': process.env.VERCEL_URL,
                'X-Title': 'Keyword Analyzer'
            },
            body: JSON.stringify({
                model: 'mistralai/mistral-7b-instruct',
                messages: [{
                    role: 'system',
                    content: 'You are a keyword analyzer. Respond ONLY with "true" or "false".'
                }, {
                    role: 'user',
                    content: `Is this keyword relevant for the given topic? Answer only with true or false.
                    Topic: "${topic}"
                    Keyword: "${keyword}"`
                }],
                temperature: 0.1,
                max_tokens: 5
            }),
            timeout: 30000 // 30 second timeout
        });

        if (!response.ok) {
            throw new Error(`API error: ${response.status}`);
        }

        const data = await response.json();
        return {
            keyword,
            matchType,
            status: data.choices?.[0]?.message?.content?.toLowerCase().trim() === 'true'
        };

    } catch (error) {
        if (retryCount < MAX_RETRIES) {
            console.log(`Retrying keyword "${keyword}" (attempt ${retryCount + 1})`);
            await new Promise(resolve => setTimeout(resolve, 5000 * (retryCount + 1)));
            return analyzeKeyword(keyword, matchType, topic, retryCount + 1);
        }
        throw error;
    }
}

// Modified bulk analysis endpoint
app.post('/api/analyze-bulk', upload.single('file'), async (req, res) => {
    try {
        if (isProcessing) {
            return res.status(409).json({ error: 'Analysis already in progress' });
        }

        // Validate API key first
        if (!process.env.OPENROUTER_API_KEY) {
            throw new Error('API key is not configured');
        }

        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }
        if (!req.body.topic) {
            return res.status(400).json({ error: 'Topic is required' });
        }

        console.log('Starting bulk analysis...');

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.getWorksheet(1);
        
        const keywords = [];
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) {
                const keyword = row.getCell(1).value?.toString().trim();
                if (keyword) {
                    keywords.push({
                        keyword,
                        matchType: row.getCell(2).value?.toString().trim() || 'Broad'
                    });
                }
            }
        });

        console.log(`Found ${keywords.length} keywords to process`);

        isProcessing = true;
        let errorCount = 0;
        analysisResults = [];

        // Process keywords with error handling and pausing
        const processKeywords = async () => {
            for (let i = 0; i < keywords.length; i += BATCH_SIZE) {
                const batch = keywords.slice(i, Math.min(i + BATCH_SIZE, keywords.length));
                
                for (const { keyword, matchType } of batch) {
                    try {
                        const result = await analyzeKeyword(keyword, matchType, req.body.topic);
                        analysisResults.push(result);
                        errorCount = 0; // Reset error count on success

                        sendProgressUpdate({
                            type: 'progress',
                            processed: analysisResults.length,
                            total: keywords.length,
                            percentComplete: Math.round((analysisResults.length / keywords.length) * 100)
                        });

                        // Delay between keywords
                        await new Promise(resolve => setTimeout(resolve, DELAY_BETWEEN_KEYWORDS));

                    } catch (error) {
                        console.error(`Error processing "${keyword}":`, error);
                        errorCount++;

                        if (errorCount >= MAX_ERRORS_BEFORE_PAUSE) {
                            console.log('Too many errors, pausing processing...');
                            await new Promise(resolve => setTimeout(resolve, ERROR_PAUSE_DURATION));
                            errorCount = 0;
                        }
                    }
                }

                // Delay between batches
                if (i + BATCH_SIZE < keywords.length) {
                    await new Promise(resolve => setTimeout(resolve, DELAY_BETWEEN_BATCHES));
                }
            }

            isProcessing = false;
            sendProgressUpdate({
                type: 'complete',
                processed: keywords.length,
                total: keywords.length,
                message: 'Analysis completed successfully!'
            });
        };

        // Start processing
        res.json({ success: true, message: 'Analysis started' });
        processKeywords().catch(error => {
            isProcessing = false;
            sendProgressUpdate({
                type: 'error',
                message: error.message
            });
        });

    } catch (error) {
        isProcessing = false;
        res.status(500).json({ error: error.message });
    }
});

// Add download endpoint
app.get('/api/download-results', async (req, res) => {
    try {
        if (!analysisResults || analysisResults.length === 0) {
            return res.status(404).json({ error: 'No results available for download' });
        }

        // Create a new workbook
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Analysis Results');

        // Add headers
        worksheet.addRow(['Keyword', 'Match Type', 'Status']);

        // Add data
        analysisResults.forEach(result => {
            worksheet.addRow([
                result.keyword,
                result.matchType,
                result.status === true ? 'Relevant' : 
                result.status === false ? 'Not Relevant' : 
                'Error'
            ]);
        });

        // Style the headers
        const headerRow = worksheet.getRow(1);
        headerRow.font = { bold: true };
        headerRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
        };

        // Auto-fit columns
        worksheet.columns.forEach(column => {
            column.width = Math.max(
                Math.max(...column.values.map(v => v ? v.toString().length : 0)),
                column.header.length
            ) + 2;
        });

        // Set response headers
        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        res.setHeader(
            'Content-Disposition',
            'attachment; filename=keyword-analysis-results.xlsx'
        );

        // Write to response
        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error('Download error:', error);
        res.status(500).json({
            error: 'Download failed',
            message: error.message
        });
    }
});

// Root path handler
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, '../public/index.html'));
});

// 404 handler
app.use((req, res) => {
    res.status(404).json({ error: 'Not Found' });
});

// Error handler
app.use((err, req, res, next) => {
    console.error(err.stack);
    res.status(500).json({ 
        error: 'Internal Server Error',
        message: process.env.NODE_ENV === 'production' ? 'Something went wrong' : err.message
    });
});

// Export the Express API
module.exports = app;