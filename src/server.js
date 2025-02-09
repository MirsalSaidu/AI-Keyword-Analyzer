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

// Global state
const clients = new Set();
let analysisResults = [];

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../public')));

// SSE endpoint
app.get('/api/progress', (req, res) => {
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.setHeader('Access-Control-Allow-Origin', '*');

    // Send initial connection message
    res.write('data: {"type": "connected"}\n\n');

    // Add client to Set
    clients.add(res);

    // Remove client on connection close
    req.on('close', () => {
        clients.delete(res);
    });
});

function sendToClients(data) {
    clients.forEach(client => {
        try {
            client.write(`data: ${JSON.stringify(data)}\n\n`);
        } catch (error) {
            console.error('Error sending to client:', error);
            clients.delete(client);
        }
    });
}

async function analyzeKeyword(keyword, matchType, topic) {
    try {
        sendToClients({
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
            })
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
        console.error(`Error analyzing "${keyword}":`, error);
        throw error;
    }
}

app.post('/api/analyze', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
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

        // Send initial response
        res.json({ 
            success: true, 
            message: 'Analysis started',
            totalKeywords: keywords.length 
        });

        // Process keywords
        analysisResults = [];
        let processed = 0;

        for (let i = 0; i < keywords.length; i += BATCH_SIZE) {
            const batch = keywords.slice(i, Math.min(i + BATCH_SIZE, keywords.length));
            
            for (const { keyword, matchType } of batch) {
                try {
                    const result = await analyzeKeyword(keyword, matchType, req.body.topic);
                    analysisResults.push(result);
                    processed++;

                    sendToClients({
                        type: 'progress',
                        processed,
                        total: keywords.length,
                        percentComplete: Math.round((processed / keywords.length) * 100)
                    });

                } catch (error) {
                    console.error(`Error processing "${keyword}":`, error);
                    analysisResults.push({ 
                        keyword, 
                        matchType, 
                        status: 'error',
                        error: error.message 
                    });
                    processed++;
                }

                await new Promise(resolve => setTimeout(resolve, DELAY_BETWEEN_KEYWORDS));
            }
        }

        sendToClients({
            type: 'complete',
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

app.get('/api/download', async (req, res) => {
    try {
        if (analysisResults.length === 0) {
            return res.status(404).json({ error: 'No results available' });
        }

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Analysis Results');

        worksheet.addRow(['Keyword', 'Match Type', 'Status']);
        analysisResults.forEach(result => {
            worksheet.addRow([
                result.keyword,
                result.matchType,
                result.status === true ? 'Relevant' : 
                result.status === false ? 'Not Relevant' : 
                'Error'
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