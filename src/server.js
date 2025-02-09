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

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../public')));

// Constants
const BATCH_SIZE = 5;
const DELAY_BETWEEN_BATCHES = 500;
const clients = new Map();

// Modified SSE setup with heartbeat
app.get('/api/analysis-progress', (req, res) => {
    const clientId = Date.now().toString(36) + Math.random().toString(36).substr(2);
    
    res.writeHead(200, {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Access-Control-Allow-Origin': '*'
    });

    // Send initial connection message
    res.write('data: {"type": "connected"}\n\n');

    // Set up heartbeat to keep connection alive
    const heartbeat = setInterval(() => {
        if (clients.has(clientId)) {
            res.write('data: {"type": "heartbeat"}\n\n');
        }
    }, 30000); // Send heartbeat every 30 seconds

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

// Keyword analysis function
async function analyzeKeyword(keyword, matchType, topic) {
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
                'HTTP-Referer': process.env.VERCEL_URL || 'http://localhost:3000',
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
        const result = {
            keyword,
            matchType,
            status: data.choices?.[0]?.message?.content?.toLowerCase().trim() === 'true'
        };

        sendProgressUpdate({
            type: 'completed',
            keyword,
            status: 'Done'
        });

        return result;

    } catch (error) {
        console.error(`Error analyzing "${keyword}":`, error);
        sendProgressUpdate({
            type: 'error',
            keyword,
            status: 'Error'
        });
        return { keyword, matchType, status: 'error' };
    }
}

// Main analysis endpoint
app.post('/api/analyze-bulk', upload.single('file'), async (req, res) => {
    try {
        // Validate input
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }
        if (!req.body.topic) {
            return res.status(400).json({ error: 'Topic is required' });
        }

        // Process Excel file
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

        // Initialize progress
        sendProgressUpdate({
            type: 'start',
            total: keywords.length,
            message: 'Starting analysis...'
        });

        // Process keywords in batches
        const results = [];
        let processedCount = 0;

        for (let i = 0; i < keywords.length; i += BATCH_SIZE) {
            const batch = keywords.slice(i, Math.min(i + BATCH_SIZE, keywords.length));
            
            const batchResults = await Promise.all(
                batch.map(({ keyword, matchType }) => 
                    analyzeKeyword(keyword, matchType, req.body.topic)
                )
            );
            
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

        // Send completion status
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
        console.error('Analysis error:', error);
        sendProgressUpdate({
            type: 'error',
            message: error.message
        });
        res.status(500).json({ 
            error: 'Analysis failed', 
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