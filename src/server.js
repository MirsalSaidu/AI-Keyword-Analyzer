const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const fetch = require('node-fetch');
const cors = require('cors');
require('dotenv').config();

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Enable CORS for all routes
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

let analysisResults = [];
const clients = new Map();

const BATCH_SIZE = 5;
const DELAY_BETWEEN_BATCHES = 500;

// For Vercel, we'll use a different approach for SSE
const clientsMap = new Map();

// Helper function to generate unique client IDs
const generateClientId = () => Date.now().toString(36) + Math.random().toString(36).substr(2);

// Modified SSE endpoint
app.get('/api/analysis-progress', (req, res) => {
    const clientId = generateClientId();
    res.writeHead(200, {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Access-Control-Allow-Origin': '*'
    });

    clients.set(clientId, res);

    req.on('close', () => {
        clients.delete(clientId);
    });
});

// Modified progress update function
function sendProgressUpdate(data) {
    const message = `data: ${JSON.stringify(data)}\n\n`;
    clients.forEach(client => {
        client.write(message);
    });
}

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
                'HTTP-Referer': 'http://localhost:3000',
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
        const content = data.choices?.[0]?.message?.content?.toLowerCase().trim();
        
        sendProgressUpdate({
            type: 'completed',
            keyword,
            status: 'Done'
        });

        return {
            keyword,
            matchType,
            status: content === 'true' ? 'true' : 'false'
        };

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

app.post('/api/analyze-bulk', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        const { topic } = req.body;
        if (!topic) {
            return res.status(400).json({ error: 'Topic is required' });
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.getWorksheet(1);
        
        const keywords = [];
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) {
                const keyword = row.getCell(1).value;
                const matchType = row.getCell(2).value || 'Broad';
                if (keyword) {
                    keywords.push({
                        keyword: keyword.toString().trim(),
                        matchType: matchType.toString().trim()
                    });
                }
            }
        });

        // Initial progress update
        sendProgressUpdate({
            type: 'start',
            total: keywords.length,
            message: 'Starting analysis...'
        });

        const results = [];
        let processedCount = 0;

        // Process in optimized batches
        for (let i = 0; i < keywords.length; i += BATCH_SIZE) {
            const batch = keywords.slice(i, Math.min(i + BATCH_SIZE, keywords.length));
            
            // Process batch concurrently
            const batchResults = await Promise.all(
                batch.map(({ keyword, matchType }) => 
                    analyzeKeyword(keyword, matchType, topic)
                )
            );
            
            results.push(...batchResults);
            processedCount += batch.length;

            // Update progress
            sendProgressUpdate({
                type: 'progress',
                processed: processedCount,
                total: keywords.length,
                percentComplete: Math.round((processedCount / keywords.length) * 100)
            });

            // Add small delay between batches to prevent rate limiting
            if (i + BATCH_SIZE < keywords.length) {
                await new Promise(resolve => setTimeout(resolve, DELAY_BETWEEN_BATCHES));
            }
        }

        // Store results
        analysisResults = results;

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
            totalKeywords: results.length,
            processedCount: processedCount
        });

    } catch (error) {
        console.error('Analysis error:', error);
        sendProgressUpdate({
            type: 'error',
            message: error.message
        });
        res.status(500).json({ error: 'Analysis failed', message: error.message });
    }
});

// Download results endpoint
app.get('/api/download-results', async (req, res) => {
    try {
        if (!analysisResults || analysisResults.length === 0) {
            return res.status(404).json({ error: 'No analysis results available' });
        }

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Keyword Analysis');

        // Add headers
        worksheet.columns = [
            { header: 'Keyword', key: 'keyword', width: 30 },
            { header: 'Match Type', key: 'matchType', width: 15 },
            { header: 'Status', key: 'status', width: 15 }
        ];

        // Add data
        worksheet.addRows(analysisResults);

        // Style header row
        worksheet.getRow(1).font = { bold: true };
        worksheet.getRow(1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
        };

        // Set response headers
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=keyword-analysis.xlsx');

        // Write to response
        await workbook.xlsx.write(res);

    } catch (error) {
        console.error('Download error:', error);
        res.status(500).json({ error: 'Error generating download file' });
    }
});

// Health check endpoint
app.get('/api/health', (req, res) => {
    res.json({ status: 'ok' });
});

// Update the port configuration for Vercel
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
    console.log('OpenRouter API Key:', process.env.OPENROUTER_API_KEY ? 'Configured' : 'Missing');
});

// Export the Express API
module.exports = app;
