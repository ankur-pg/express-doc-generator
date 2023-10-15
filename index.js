import express from 'express';
import fs from 'fs';
import path from 'path';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();

app.get('/generate-doc', (req, res) => {
    const templatePath = path.join(__dirname, 'template.docx');

    fs.promises.readFile(templatePath, 'binary').then(content => {
        const zip = new PizZip(content);
        const doc = new Docxtemplater().loadZip(zip);

        const data = {
            name1: 'John Doe',
            age1: 30,
            name2: 'Jane Smith',
            age2: 25
            // ... add more as needed
        };

        doc.setData(data);

        try {
            doc.render();
        } catch (error) {
            console.error('Error during rendering:', error);
            res.status(500).send('Document generation failed.');
            return;
        }

        const docBuffer = doc.getZip().generate({ type: 'nodebuffer' });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', 'attachment; filename=GeneratedDocument.docx');
        res.send(docBuffer);
    }).catch(err => {
        console.error('Error reading template:', err);
        res.status(500).send('Failed to read template.');
    });
});

app.listen(3000, () => {
    console.log('Server started on http://localhost:3000');
});
