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
const BATCH_SIZE = 10; // Increased batch size
const DELAY_BETWEEN_BATCHES = 300; // Reduced delay (ms)
const REQUEST_TIMEOUT = 300000; // 5 minutes timeout
const KEEP_ALIVE_TIMEOUT = 305000; // Slightly longer than request timeout

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

// Modified SSE setup with longer timeouts
app.get('/api/analysis-progress', (req, res) => {
    const clientId = Date.now().toString(36) + Math.random().toString(36).substr(2);
    
    req.socket.setTimeout(KEEP_ALIVE_TIMEOUT);
    res.writeHead(200, {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Access-Control-Allow-Origin': '*',
        'X-Accel-Buffering': 'no' // Disable proxy buffering
    });

    // Send initial connection message
    res.write('data: {"type": "connected"}\n\n');

    // Heartbeat every 30 seconds
    const heartbeat = setInterval(() => {
        if (clients.has(clientId)) {
            res.write('data: {"type": "heartbeat"}\n\n');
        }
    }, 30000);

    clients.set(clientId, res);

    req.on('close', () => {
        clients.delete(clientId);
        clearInterval(heartbeat);
    });
});

// Improved sendProgressUpdate function
function sendProgressUpdate(data) {
    const message = `data: ${JSON.stringify(data)}\n\n`;
    clients.forEach((client, clientId) => {
        try {
            client.write(message);
        } catch (error) {
            console.error(`Error sending to client ${clientId}:`, error);
            clients.delete(clientId);
        }
    });
}

// Modified analyze keyword function with better error handling and logging
async function analyzeKeyword(keyword, matchType, topic) {
    try {
        console.log(`Processing keyword: "${keyword}" for topic: "${topic}"`);
        
        sendProgressUpdate({
            type: 'processing',
            keyword,
            status: 'Processing'
        });

        const controller = new AbortController();
        const timeout = setTimeout(() => controller.abort(), 60000);

        const apiUrl = 'https://openrouter.ai/api/v1/chat/completions';
        const headers = {
            'Authorization': `Bearer ${process.env.OPENROUTER_API_KEY}`,
            'Content-Type': 'application/json',
            'HTTP-Referer': process.env.VERCEL_URL || 'https://your-vercel-domain.vercel.app',
            'X-Title': 'Keyword Analyzer'
        };

        console.log('Sending request to OpenRouter API...');
        
        const response = await fetch(apiUrl, {
            method: 'POST',
            headers: headers,
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
            signal: controller.signal
        });

        clearTimeout(timeout);

        if (!response.ok) {
            const errorText = await response.text();
            console.error('API Response Error:', {
                status: response.status,
                statusText: response.statusText,
                body: errorText
            });
            throw new Error(`API error: ${response.status} - ${errorText}`);
        }

        const data = await response.json();
        console.log('API Response:', data);

        if (!data.choices?.[0]?.message?.content) {
            throw new Error('Invalid API response format');
        }

        const content = data.choices[0].message.content.toLowerCase().trim();
        console.log(`Analysis result for "${keyword}": ${content}`);

        const result = {
            keyword,
            matchType,
            status: content === 'true'
        };

        sendProgressUpdate({
            type: 'completed',
            keyword,
            status: 'Done'
        });

        return result;

    } catch (error) {
        console.error(`Error analyzing "${keyword}":`, error);
        
        // Check if it's an API key error
        if (error.message.includes('401') || error.message.includes('unauthorized')) {
            console.error('API Key authentication failed');
            sendProgressUpdate({
                type: 'error',
                message: 'API authentication failed. Please check your API key.'
            });
        }
        
        sendProgressUpdate({
            type: 'error',
            keyword,
            status: 'Error',
            message: error.message
        });
        
        return { 
            keyword, 
            matchType, 
            status: 'error',
            error: error.message 
        };
    }
}

// Modified bulk analysis endpoint with better error handling
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
        console.log('File size:', req.file.size);
        console.log('Topic:', req.body.topic);

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

        sendProgressUpdate({
            type: 'start',
            total: keywords.length,
            message: 'Starting analysis...'
        });

        const results = [];
        let processedCount = 0;

        // Process in larger chunks with controlled concurrency
        for (let i = 0; i < keywords.length; i += BATCH_SIZE) {
            const batch = keywords.slice(i, Math.min(i + BATCH_SIZE, keywords.length));
            
            const batchPromises = batch.map(({ keyword, matchType }) => 
                analyzeKeyword(keyword, matchType, req.body.topic)
                    .catch(error => ({ keyword, matchType, status: 'error', error: error.message }))
            );

            const batchResults = await Promise.all(batchPromises);
            results.push(...batchResults);
            processedCount += batch.length;

            sendProgressUpdate({
                type: 'progress',
                processed: processedCount,
                total: keywords.length,
                percentComplete: Math.round((processedCount / keywords.length) * 100)
            });

            if (i + BATCH_SIZE < keywords.length) {
                await new Promise(resolve => setTimeout(resolve, DELAY_BETWEEN_BATCHES));
            }
        }

        sendProgressUpdate({
            type: 'complete',
            processed: keywords.length,
            total: keywords.length,
            message: 'Analysis completed successfully!'
        });

        res.json({
            success: true,
            message: 'Analysis completed successfully',
            totalKeywords: results.length
        });

    } catch (error) {
        console.error('Bulk analysis error:', error);
        sendProgressUpdate({
            type: 'error',
            message: error.message
        });
        res.status(500).json({ 
            error: 'Analysis failed', 
            message: process.env.NODE_ENV === 'production' 
                ? 'An error occurred during analysis' 
                : error.message 
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