const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const cors = require('cors');
const PDFToWordConverter = require('./converter');

const app = express();
const upload = multer({ dest: 'uploads/' });
const converter = new PDFToWordConverter();

app.use(cors());
app.use(express.static('public'));

const progressMap = new Map();

app.post('/convert', upload.single('pdf'), async (req, res) => {
    if (!req.file) {
        return res.status(400).send('No file uploaded.');
    }

    const jobId = Math.random().toString(36).substring(7);
    const inputPath = req.file.path;
    const outputFileName = `${req.file.originalname.replace('.pdf', '')}.docx`;
    const outputPath = path.join('outputs', outputFileName);

    if (!fs.existsSync('outputs')) {
        fs.mkdirSync('outputs');
    }

    // Start conversion in background
    progressMap.set(jobId, { percent: 0, status: 'بدء المعالجة...' });

    // Respond with Job ID immediately
    res.json({ success: true, jobId });

    (async () => {
        try {
            await converter.convert(inputPath, outputPath, (percent) => {
                // Don't send 100% until we have the downloadUrl ready in the next step
                if (percent < 100) {
                    progressMap.set(jobId, { percent, status: `جاري التحويل... ${percent}%` });
                }
            });
            // Final update with the URL
            progressMap.set(jobId, {
                percent: 100,
                status: 'اكتمل بنجاح!',
                downloadUrl: `/download/${outputFileName}`,
                completed: true
            });
        } catch (error) {
            console.error('Conversion error:', error);
            progressMap.set(jobId, { percent: 0, status: 'Error', message: error.message });
        }
    })();
});

app.get('/progress/:jobId', (req, res) => {
    const status = progressMap.get(req.params.jobId);
    if (!status) return res.status(404).json({ message: 'Job not found' });
    res.json(status);
});

app.get('/download/:filename', (req, res) => {
    const filePath = path.join(__dirname, 'outputs', req.params.filename);
    if (fs.existsSync(filePath)) {
        res.download(filePath);
    } else {
        res.status(404).send('File not found.');
    }
});

const PORT = 3000;
app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
