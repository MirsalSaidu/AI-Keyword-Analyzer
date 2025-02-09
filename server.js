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

// Constants for timing and batches
const BATCH_SIZE = 10;
const DELAY_BETWEEN_BATCHES = 300;
const REQUEST_TIMEOUT = 300000; // 5 minutes timeout
const KEEP_ALIVE_TIMEOUT = 305000; // Slightly longer than request timeout
const SSE_RETRY_TIMEOUT = 15000; // 15 seconds retry timeout

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(express.static(path.join(__dirname, '../public')));

// Configure request timeouts
app.use((req, res, next) => {
    res.setTimeout(REQUEST_TIMEOUT, () => {
        res.status(408).send('Request timeout');
    });
    next();
});

// Constants
const clients = new Map();

// Add API key validation
if (!process.env.OPENROUTER_API_KEY) {
    console.error('OPENROUTER_API_KEY is not set in environment variables');
}

// Modified SSE setup with better connection handling
app.get('/api/analysis-progress', (req, res) => {
    const clientId = Date.now().toString(36) + Math.random().toString(36).substr(2);
    
    // Set headers for better SSE handling
    res.writeHead(200, {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache, no-transform',
        'Connection': 'keep-alive',
        'X-Accel-Buffering': 'no',
        'Access-Control-Allow-Origin': '*',
        'retry': SSE_RETRY_TIMEOUT
    });

    // Send initial connection message with retry information
    res.write(`retry: ${SSE_RETRY_TIMEOUT}\n`);
    res.write('data: {"type": "connected"}\n\n');

    // More frequent heartbeat (every 15 seconds)
    const heartbeat = setInterval(() => {
        if (clients.has(clientId)) {
            try {
                res.write('data: {"type": "heartbeat"}\n\n');
            } catch (error) {
                console.error('Heartbeat error:', error);
                cleanup();
            }
        }
    }, 15000);

    // Keep connection alive
    req.socket.setKeepAlive(true);
    req.socket.setTimeout(0);

    // Cleanup function
    const cleanup = () => {
        clearInterval(heartbeat);
        clients.delete(clientId);
    };

    // Store client
    clients.set(clientId, {
        res,
        timestamp: Date.now(),
        cleanup
    });

    // Handle client disconnect
    req.on('close', cleanup);
    req.on('end', cleanup);
    req.on('error', cleanup);
});

// Improved sendProgressUpdate function
function sendProgressUpdate(data) {
    const now = Date.now();
    const message = `data: ${JSON.stringify({ ...data, timestamp: now })}\n\n`;
    
    clients.forEach((client, clientId) => {
        try {
            client.res.write(message);
        } catch (error) {
            console.error(`Error sending to client ${clientId}:`, error);
            client.cleanup();
        }
    });
}

// Modified analyze keyword function with better error handling and logging
async function analyzeKeyword(keyword, matchType, topic) {
    try {
        console.log(`Analyzing keyword: "${keyword}"`);
        
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
                'HTTP-Referer': process.env.VERCEL_URL || 'https://keyword-analyzer.vercel.app',
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

        if (!response.ok) {
            throw new Error(`API error: ${response.status}`);
        }

        const data = await response.json();
        console.log(`API response for "${keyword}":`, data);

        sendProgressUpdate({
            type: 'completed',
            keyword,
            status: 'Done'
        });

        return {
            keyword,
            matchType,
            status: data.choices?.[0]?.message?.content?.toLowerCase().trim() === 'true'
        };

    } catch (error) {
        console.error(`Error analyzing "${keyword}":`, error);
        sendProgressUpdate({
            type: 'error',
            keyword,
            status: 'Error'
        });
        throw error; // Re-throw to handle in the batch processing
    }
}

// Modified bulk analysis endpoint
app.post('/api/analyze-bulk', upload.single('file'), async (req, res) => {
    try {
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

        // Send initial total count immediately
        sendProgressUpdate({
            type: 'start',
            total: keywords.length,
            message: 'Starting analysis...'
        });

        // Send immediate response to prevent timeout
        res.json({
            success: true,
            message: 'Analysis started',
            totalKeywords: keywords.length
        });

        // Process keywords with better error handling
        const processKeywords = async () => {
            const results = [];
            let processedCount = 0;
            let consecutiveErrors = 0;
            const MAX_CONSECUTIVE_ERRORS = 3;

            for (let i = 0; i < keywords.length; i += BATCH_SIZE) {
                const batch = keywords.slice(i, Math.min(i + BATCH_SIZE, keywords.length));

                for (const { keyword, matchType } of batch) {
                    try {
                        const result = await analyzeKeyword(keyword, matchType, req.body.topic);
                        results.push(result);
                        processedCount++;
                        consecutiveErrors = 0; // Reset error counter on success

                        sendProgressUpdate({
                            type: 'progress',
                            processed: processedCount,
                            total: keywords.length,
                            percentComplete: Math.round((processedCount / keywords.length) * 100)
                        });

                    } catch (error) {
                        console.error(`Error processing keyword "${keyword}":`, error);
                        consecutiveErrors++;
                        
                        // Add error result
                        results.push({ 
                            keyword, 
                            matchType, 
                            status: 'error',
                            error: error.message 
                        });
                        
                        processedCount++;

                        // If too many consecutive errors, pause processing
                        if (consecutiveErrors >= MAX_CONSECUTIVE_ERRORS) {
                            console.log('Too many consecutive errors, pausing for 30 seconds...');
                            await new Promise(resolve => setTimeout(resolve, 30000));
                            consecutiveErrors = 0;
                        }
                    }

                    // Add small delay between individual keywords
                    await new Promise(resolve => setTimeout(resolve, 100));
                }

                // Delay between batches
                if (i + BATCH_SIZE < keywords.length) {
                    await new Promise(resolve => setTimeout(resolve, DELAY_BETWEEN_BATCHES));
                }
            }

            // Send completion update
            sendProgressUpdate({
                type: 'complete',
                processed: keywords.length,
                total: keywords.length,
                message: 'Analysis completed successfully!'
            });
        };

        // Start processing with error handling
        processKeywords().catch(error => {
            console.error('Background processing error:', error);
            sendProgressUpdate({
                type: 'error',
                message: 'Processing failed: ' + error.message
            });
        });

    } catch (error) {
        console.error('Analysis setup error:', error);
        res.status(500).json({ 
            error: 'Analysis failed to start', 
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