const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const fetch = require('node-fetch');
const cors = require('cors');
const path = require('path');
require('dotenv').config();

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Constants
const BATCH_SIZE = 2;
const DELAY_BETWEEN_KEYWORDS = 1000;
const MAX_RETRIES = 3;
const SSE_TIMEOUT = 8 * 60 * 60 * 1000; // 8 hours
const HEARTBEAT_INTERVAL = 30000; // 30 seconds

// Global state
const clients = new Map();
let analysisResults = [];

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../public')));

// Status endpoint for polling
app.get('/api/status', (req, res) => {
    res.json({
        ...processingState,
        timestamp: Date.now()
    });
});

// Modified SSE endpoint with extended timeout
app.get('/api/analysis-progress', (req, res) => {
    // Increase timeouts
    req.socket.setTimeout(SSE_TIMEOUT);
    res.socket.setTimeout(SSE_TIMEOUT);

    // Set headers for stable connection
    res.writeHead(200, {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache, no-transform',
        'Connection': 'keep-alive',
        'X-Accel-Buffering': 'no',
        'Access-Control-Allow-Origin': '*',
        'Keep-Alive': `timeout=${SSE_TIMEOUT}`
    });

    const clientId = Date.now().toString();

    // Send initial connection message
    res.write('data: {"type": "connected"}\n\n');

    // Setup heartbeat interval
    const heartbeat = setInterval(() => {
        try {
            res.write('data: {"type": "heartbeat"}\n\n');
        } catch (error) {
            console.error('Heartbeat error:', error);
            cleanup();
        }
    }, HEARTBEAT_INTERVAL);

    // Cleanup function
    const cleanup = () => {
        clearInterval(heartbeat);
        clients.delete(clientId);
    };

    // Store client with its cleanup function
    clients.set(clientId, {
        res,
        cleanup,
        startTime: Date.now()
    });

    // Handle client disconnect
    req.on('close', cleanup);
    req.on('error', cleanup);
});

// Modified sendToClients function with error handling
function sendToClients(data) {
    const now = Date.now();
    
    clients.forEach((client, clientId) => {
        try {
            // Check if client connection is still within timeout
            if (now - client.startTime < SSE_TIMEOUT) {
                client.res.write(`data: ${JSON.stringify({
                    ...data,
                    timestamp: now
                })}\n\n`);
            } else {
                console.log(`Client ${clientId} timed out, cleaning up`);
                client.cleanup();
            }
        } catch (error) {
            console.error(`Error sending to client ${clientId}:`, error);
            client.cleanup();
        }
    });
}

// Analyze keyword function with retries
async function analyzeKeyword(keyword, matchType, topic, retryCount = 0) {
    try {
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
            })
        });

        if (response.status === 429) {
            if (retryCount < MAX_RETRIES) {
                await new Promise(resolve => setTimeout(resolve, 5000 * (retryCount + 1)));
                return analyzeKeyword(keyword, matchType, topic, retryCount + 1);
            }
            throw new Error('Rate limit exceeded');
        }

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
        console.error(`Error analyzing "${keyword}":`, error);
        throw error;
    }
}

// Modified bulk analysis endpoint with better error handling
app.post('/api/analyze-bulk', upload.single('file'), async (req, res) => {
    try {
        if (processingState.isProcessing) {
            return res.status(409).json({ error: 'Analysis already in progress' });
        }

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

        if (keywords.length === 0) {
            return res.status(400).json({ error: 'No keywords found in file' });
        }

        // Initialize processing state
        processingState = {
            isProcessing: true,
            currentBatch: 0,
            totalBatches: Math.ceil(keywords.length / BATCH_SIZE),
            processedCount: 0,
            totalKeywords: keywords.length,
            results: [],
            errors: [],
            lastUpdateTime: Date.now()
        };

        // Send initial response
        res.json({ 
            success: true, 
            message: 'Analysis started',
            totalKeywords: keywords.length
        });

        // Process keywords with better error handling
        for (let i = 0; i < keywords.length; i += BATCH_SIZE) {
            const batch = keywords.slice(i, Math.min(i + BATCH_SIZE, keywords.length));
            
            for (const { keyword, matchType } of batch) {
                try {
                    const result = await analyzeKeyword(keyword, matchType, req.body.topic);
                    analysisResults.push(result);

                    // Send progress update
                    sendToClients({
                        type: 'progress',
                        processed: analysisResults.length,
                        total: keywords.length,
                        currentKeyword: keyword,
                        percentComplete: Math.round((analysisResults.length / keywords.length) * 100)
                    });

                } catch (error) {
                    console.error(`Error processing "${keyword}":`, error);
                    
                    // Handle rate limiting
                    if (error.message.includes('429')) {
                        console.log('Rate limit hit, pausing...');
                        await new Promise(resolve => setTimeout(resolve, 60000)); // 1 minute pause
                        i -= BATCH_SIZE; // Retry this batch
                        continue;
                    }

                    analysisResults.push({
                        keyword,
                        matchType,
                        status: 'error',
                        error: error.message
                    });
                }

                await new Promise(resolve => setTimeout(resolve, DELAY_BETWEEN_KEYWORDS));
            }
        }

        // Send completion message
        sendToClients({
            type: 'complete',
            processed: keywords.length,
            total: keywords.length,
            message: 'Analysis completed successfully!'
        });

    } catch (error) {
        console.error('Analysis error:', error);
        sendToClients({
            type: 'error',
            message: error.message
        });
    }
});

// Download results endpoint
app.get('/api/download-results', async (req, res) => {
    try {
        if (processingState.results.length === 0) {
            return res.status(404).json({ error: 'No results available' });
        }

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Analysis Results');

        worksheet.addRow(['Keyword', 'Match Type', 'Status']);
        processingState.results.forEach(result => {
            worksheet.addRow([
                result.keyword,
                result.matchType,
                result.status === true ? 'Relevant' : 'Not Relevant'
            ]);
        });

        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        res.setHeader(
            'Content-Disposition',
            'attachment; filename=keyword-analysis-results.xlsx'
        );

        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error('Download error:', error);
        res.status(500).json({ error: 'Download failed' });
    }
});

module.exports = app;