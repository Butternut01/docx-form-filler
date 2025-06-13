const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');
const Docxtemplater = require('docxtemplater');
const PizZip = require('pizzip');
const mongoose = require('mongoose');

// App setup
const app = express();
const port = 3000;

// MongoDB setup
mongoose.connect('mongodb://localhost:27017/documentsDB', {
    useNewUrlParser: true,
    useUnifiedTopology: true,
})
.then(() => console.log(' Connected to MongoDB'))
.catch((err) => console.error(' MongoDB connection error:', err));

// Mongoose Schema
const generatedFileSchema = new mongoose.Schema({
    filename: String,
    data: Buffer,
    contentType: String,
    createdAt: { type: Date, default: Date.now }
});
const GeneratedFile = mongoose.model('GeneratedFile', generatedFileSchema);

// Middleware
app.use(bodyParser.json());
app.use(express.static('public'));

// Serve HTML form
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Generate DOCX and Save to MongoDB
app.post('/generate-docx', async (req, res) => {
    try {
        const templatePath = path.join(__dirname, 'template.docx');
        const content = fs.readFileSync(templatePath, 'binary');
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        const data = {
            responsible: req.body.responsible,
            meetingTopic: req.body.meetingTopic,
            needLaptop: req.body.needLaptop,
            weekday: req.body.weekday,
            date: req.body.date,
            startTime: req.body.startTime,
            endTime: req.body.endTime,
            recording: req.body.recording,
            participants: req.body.participants.map((p, i) => ({
                num: i + 1,
                name: p.name,
                position: p.position,
                phone: p.phone,
                email: p.email,
                connection: p.connection === 'ссылка' ? 'ссылка' : 'терминал',
                location: p.location
            })),
            emptyRows: Array(Math.max(0, 8 - req.body.participants.length)).fill({})
        };

        doc.render(data);

        const buf = doc.getZip().generate({
            type: 'nodebuffer',
            compression: 'DEFLATE'
        });

        const filename = 'Заявка_на_ВКС.docx';

        // Save to MongoDB
        const newFile = new GeneratedFile({
            filename,
            data: buf,
            contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });

        await newFile.save();
        console.log(' File saved to MongoDB');

        // Send file to user
        const encodedName = encodeURIComponent(filename);
        res.setHeader('Content-Type', newFile.contentType);
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodedName}`);
        res.send(buf);

    } catch (error) {
        console.error(' Error generating document:', error);
        res.status(500).send('Error generating document');
    }
});

app.listen(port, () => {
    console.log(` http://localhost:${port}`);
});
