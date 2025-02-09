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

// Constants for large-scale processing
const BATCH_SIZE = 2; // Very small batch size for stability
const DELAY_BETWEEN_BATCHES = 3000; // 3 seconds between batches
const DELAY_BETWEEN_KEYWORDS = 1000; // 1 second between keywords
const MAX_RETRIES = 3;
const RATE_LIMIT_PAUSE = 60000; // 1 minute pause when rate limited
const MAX_CONCURRENT_CONNECTIONS = 100;

// SSE Constants
const SSE_KEEPALIVE_INTERVAL = 3000; // 3 seconds
const SSE_TIMEOUT = 8 * 60 * 60 * 1000; // 8 hours
const PROGRESS_CHUNK_SIZE = 10; // Update progress every 10 keywords

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

// Global state
let globalProcessingState = {
    isProcessing: false,
    processedCount: 0,
    totalKeywords: 0,
    results: [],
    errors: [],
    startTime: Date.now(),
    lastUpdateTime: Date.now()
};

// Create a new connection manager
class ConnectionManager {
    constructor() {
        this.clients = new Map();
        this.setupPeriodicCleanup();
    }

    addClient(clientId, res) {
        const client = {
            res,
            startTime: Date.now(),
            lastPingTime: Date.now(),
            isAlive: true
        };

        this.clients.set(clientId, client);
        this.setupClientKeepalive(clientId);
        return client;
    }

    setupClientKeepalive(clientId) {
        const keepaliveInterval = setInterval(() => {
            const client = this.clients.get(clientId);
            if (client && client.isAlive) {
                try {
                    client.res.write(`data: ${JSON.stringify({
                        type: 'keepalive',
                        timestamp: Date.now(),
                        state: {
                            isProcessing: globalProcessingState.isProcessing,
                            processedCount: globalProcessingState.processedCount,
                            totalKeywords: globalProcessingState.totalKeywords
                        }
                    })}\n\n`);
                    client.lastPingTime = Date.now();
                } catch (error) {
                    console.error(`Keepalive error for client ${clientId}:`, error);
                    this.removeClient(clientId);
                }
            } else {
                clearInterval(keepaliveInterval);
            }
        }, SSE_KEEPALIVE_INTERVAL);

        // Auto-cleanup after timeout
        setTimeout(() => this.removeClient(clientId), SSE_TIMEOUT);
    }

    removeClient(clientId) {
        const client = this.clients.get(clientId);
        if (client) {
            client.isAlive = false;
            try {
                client.res.end();
            } catch (error) {
                console.error(`Error ending response for client ${clientId}:`, error);
            }
            this.clients.delete(clientId);
        }
    }

    broadcast(data) {
        const message = `data: ${JSON.stringify({
            ...data,
            timestamp: Date.now(),
            state: {
                isProcessing: globalProcessingState.isProcessing,
                processedCount: globalProcessingState.processedCount,
                totalKeywords: globalProcessingState.totalKeywords
            }
        })}\n\n`;

        this.clients.forEach((client, clientId) => {
            if (client.isAlive) {
                try {
                    client.res.write(message);
                } catch (error) {
                    console.error(`Broadcast error for client ${clientId}:`, error);
                    this.removeClient(clientId);
                }
            }
        });
    }

    setupPeriodicCleanup() {
        setInterval(() => {
            const now = Date.now();
            this.clients.forEach((client, clientId) => {
                if (now - client.lastPingTime > SSE_KEEPALIVE_INTERVAL * 2) {
                    console.log(`Client ${clientId} timed out`);
                    this.removeClient(clientId);
                }
            });
        }, SSE_KEEPALIVE_INTERVAL);
    }
}

const connectionManager = new ConnectionManager();

// Add API key validation
if (!process.env.OPENROUTER_API_KEY) {
    console.error('OPENROUTER_API_KEY is not set in environment variables');
}

// Modified SSE endpoint
app.get('/api/analysis-progress', (req, res) => {
    const clientId = Date.now().toString(36) + Math.random().toString(36).substr(2);
    
    res.writeHead(200, {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache, no-transform',
        'Connection': 'keep-alive',
        'X-Accel-Buffering': 'no',
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Credentials': 'true'
    });

    const client = connectionManager.addClient(clientId, res);

    // Send initial state
    if (globalProcessingState.isProcessing) {
        client.res.write(`data: ${JSON.stringify({
            type: 'reconnect-state',
            state: globalProcessingState
        })}\n\n`);
    }

    req.on('close', () => connectionManager.removeClient(clientId));
});

// Rate limiting
const rateLimiter = {
    tokens: 50,
    lastRefill: Date.now(),
    refillRate: 50, // tokens per minute
    refillInterval: 60000, // 1 minute

    async getToken() {
        const now = Date.now();
        const timePassed = now - this.lastRefill;
        
        if (timePassed >= this.refillInterval) {
            this.tokens = this.refillRate;
            this.lastRefill = now;
        }

        if (this.tokens <= 0) {
            await new Promise(resolve => setTimeout(resolve, this.refillInterval));
            return this.getToken();
        }

        this.tokens--;
        return true;
    }
};

// Modified keyword analysis with rate limiting
async function analyzeKeyword(keyword, matchType, topic, retryCount = 0) {
    try {
        await rateLimiter.getToken();

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
            console.log('Rate limited, pausing...');
            await new Promise(resolve => setTimeout(resolve, RATE_LIMIT_PAUSE));
            return analyzeKeyword(keyword, matchType, topic, retryCount);
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
        if (retryCount < MAX_RETRIES) {
            const delay = Math.pow(2, retryCount) * 5000; // Exponential backoff
            await new Promise(resolve => setTimeout(resolve, delay));
            return analyzeKeyword(keyword, matchType, topic, retryCount + 1);
        }
        throw error;
    }
}

// Modified bulk analysis endpoint
app.post('/api/analyze-bulk', upload.single('file'), async (req, res) => {
    try {
        if (globalProcessingState.isProcessing) {
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

        console.log(`Starting analysis of ${keywords.length} keywords`);

        // Initialize global state
        globalProcessingState = {
            isProcessing: true,
            processedCount: 0,
            totalKeywords: keywords.length,
            results: [],
            errors: [],
            startTime: Date.now(),
            lastUpdateTime: Date.now()
        };

        // Send immediate response
        res.json({ 
            success: true, 
            message: 'Analysis started',
            totalKeywords: keywords.length,
            estimatedTime: Math.ceil(keywords.length * 1.5) // Estimated minutes
        });

        // Process in chunks with progress tracking
        let lastProgressUpdate = 0;

        for (let i = 0; i < keywords.length; i += BATCH_SIZE) {
            const batch = keywords.slice(i, Math.min(i + BATCH_SIZE, keywords.length));
            
            for (const { keyword, matchType } of batch) {
                try {
                    const result = await analyzeKeyword(keyword, matchType, req.body.topic);
                    globalProcessingState.results.push(result);
                    globalProcessingState.processedCount++;

                    // Update progress in chunks to reduce messages
                    if (globalProcessingState.processedCount - lastProgressUpdate >= PROGRESS_CHUNK_SIZE) {
                        connectionManager.broadcast({
                            type: 'progress',
                            processed: globalProcessingState.processedCount,
                            total: globalProcessingState.totalKeywords,
                            percentComplete: Math.round((globalProcessingState.processedCount / globalProcessingState.totalKeywords) * 100),
                            estimatedTimeRemaining: Math.ceil((globalProcessingState.totalKeywords - globalProcessingState.processedCount) * 1.5)
                        });
                        lastProgressUpdate = globalProcessingState.processedCount;
                    }

                } catch (error) {
                    console.error(`Error processing "${keyword}":`, error);
                    globalProcessingState.errors.push({ keyword, error: error.message });
                    
                    if (error.message.includes('429')) {
                        await new Promise(resolve => setTimeout(resolve, RATE_LIMIT_PAUSE));
                    }
                }

                await new Promise(resolve => setTimeout(resolve, DELAY_BETWEEN_KEYWORDS));
            }

            await new Promise(resolve => setTimeout(resolve, DELAY_BETWEEN_BATCHES));
        }

        // Final update
        globalProcessingState.isProcessing = false;
        connectionManager.broadcast({
            type: 'complete',
            processed: globalProcessingState.processedCount,
            total: globalProcessingState.totalKeywords,
            message: 'Analysis completed successfully!',
            totalTime: Math.round((Date.now() - globalProcessingState.startTime) / 60000)
        });

    } catch (error) {
        console.error('Analysis error:', error);
        globalProcessingState.isProcessing = false;
        connectionManager.broadcast({
            type: 'error',
            message: error.message
        });
    }
});

// Add download endpoint
app.get('/api/download-results', async (req, res) => {
    try {
        if (!globalProcessingState.results || globalProcessingState.results.length === 0) {
            return res.status(404).json({ error: 'No results available for download' });
        }

        // Create a new workbook
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Analysis Results');

        // Add headers
        worksheet.addRow(['Keyword', 'Match Type', 'Status']);

        // Add data
        globalProcessingState.results.forEach(result => {
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