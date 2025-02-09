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

// Global state
let processingState = {
    isProcessing: false,
    processedCount: 0,
    totalKeywords: 0,
    results: [],
    errors: [],
    startTime: null,
    lastUpdateTime: null,
    currentKeyword: '',
    status: 'idle' // idle, processing, completed, error
};

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

// Status endpoint for polling
app.get('/api/status', (req, res) => {
    res.json({
        ...processingState,
        timestamp: Date.now()
    });
});

// Rate limiter
const rateLimiter = {
    tokens: 50,
    lastRefill: Date.now(),
    refillRate: 50,
    refillInterval: 60000,

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

// Keyword analysis function
async function analyzeKeyword(keyword, matchType, topic, retryCount = 0) {
    try {
        await rateLimiter.getToken();
        processingState.currentKeyword = keyword;

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
            const delay = Math.pow(2, retryCount) * 5000;
            await new Promise(resolve => setTimeout(resolve, delay));
            return analyzeKeyword(keyword, matchType, topic, retryCount + 1);
        }
        throw error;
    }
}

// Bulk analysis endpoint
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

        // Reset and initialize state
        processingState = {
            isProcessing: true,
            processedCount: 0,
            totalKeywords: keywords.length,
            results: [],
            errors: [],
            startTime: Date.now(),
            lastUpdateTime: Date.now(),
            currentKeyword: '',
            status: 'processing'
        };

        // Send immediate response
        res.json({ 
            success: true, 
            message: 'Analysis started',
            totalKeywords: keywords.length
        });

        // Process keywords
        for (let i = 0; i < keywords.length; i += BATCH_SIZE) {
            const batch = keywords.slice(i, Math.min(i + BATCH_SIZE, keywords.length));
            
            for (const { keyword, matchType } of batch) {
                try {
                    const result = await analyzeKeyword(keyword, matchType, req.body.topic);
                    processingState.results.push(result);
                    processingState.processedCount++;
                    processingState.lastUpdateTime = Date.now();

                } catch (error) {
                    console.error(`Error processing "${keyword}":`, error);
                    processingState.errors.push({ keyword, error: error.message });
                    
                    if (error.message.includes('429')) {
                        await new Promise(resolve => setTimeout(resolve, RATE_LIMIT_PAUSE));
                    }
                }

                await new Promise(resolve => setTimeout(resolve, DELAY_BETWEEN_KEYWORDS));
            }

            await new Promise(resolve => setTimeout(resolve, DELAY_BETWEEN_BATCHES));
        }

        // Mark as completed
        processingState.isProcessing = false;
        processingState.status = 'completed';
        processingState.currentKeyword = '';

    } catch (error) {
        console.error('Analysis error:', error);
        processingState.isProcessing = false;
        processingState.status = 'error';
        processingState.errors.push({ error: error.message });
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