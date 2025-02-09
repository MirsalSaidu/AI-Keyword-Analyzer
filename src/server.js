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
const BATCH_SIZE = 1; // Process one keyword at a time
const DELAY_BETWEEN_KEYWORDS = 2000; // 2 seconds between keywords
const MAX_RETRIES = 3;
const MAX_PROCESSING_TIME = 30000; // 30 seconds max per keyword

// Global state
let globalState = {
    isProcessing: false,
    processedCount: 0,
    totalKeywords: 0,
    results: [],
    errors: [],
    currentKeyword: '',
    lastUpdateTime: Date.now()
};

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../public')));

// Status endpoint for checking progress
app.get('/api/status', (req, res) => {
    res.json({
        ...globalState,
        timestamp: Date.now()
    });
});

// Analyze keyword with timeout and retries
async function analyzeKeyword(keyword, matchType, topic, retryCount = 0) {
    try {
        const controller = new AbortController();
        const timeout = setTimeout(() => controller.abort(), MAX_PROCESSING_TIME);

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
            signal: controller.signal
        });

        clearTimeout(timeout);

        if (response.status === 429) {
            console.log('Rate limit hit, waiting...');
            await new Promise(resolve => setTimeout(resolve, 5000 * (retryCount + 1)));
            if (retryCount < MAX_RETRIES) {
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
        if (error.name === 'AbortError') {
            throw new Error('Request timeout');
        }
        throw error;
    }
}

// Bulk analysis endpoint
app.post('/api/analyze-bulk', upload.single('file'), async (req, res) => {
    try {
        if (globalState.isProcessing) {
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

        // Reset global state
        globalState = {
            isProcessing: true,
            processedCount: 0,
            totalKeywords: keywords.length,
            results: [],
            errors: [],
            currentKeyword: '',
            lastUpdateTime: Date.now()
        };

        // Send initial response
        res.json({ 
            success: true, 
            message: 'Analysis started',
            totalKeywords: keywords.length
        });

        // Process keywords in background
        (async () => {
            for (const { keyword, matchType } of keywords) {
                try {
                    globalState.currentKeyword = keyword;
                    const result = await analyzeKeyword(keyword, matchType, req.body.topic);
                    globalState.results.push(result);
                    globalState.processedCount++;
                    globalState.lastUpdateTime = Date.now();

                } catch (error) {
                    console.error(`Error processing "${keyword}":`, error);
                    globalState.errors.push({ keyword, error: error.message });
                    globalState.processedCount++;
                }

                await new Promise(resolve => setTimeout(resolve, DELAY_BETWEEN_KEYWORDS));
            }

            globalState.isProcessing = false;
            globalState.currentKeyword = '';

        })().catch(error => {
            console.error('Processing error:', error);
            globalState.isProcessing = false;
            globalState.errors.push({ error: error.message });
        });

    } catch (error) {
        console.error('Setup error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Download results endpoint
app.get('/api/download-results', async (req, res) => {
    try {
        if (globalState.results.length === 0) {
            return res.status(404).json({ error: 'No results available' });
        }

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Analysis Results');

        worksheet.addRow(['Keyword', 'Match Type', 'Status']);
        globalState.results.forEach(result => {
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