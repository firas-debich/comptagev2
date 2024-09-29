const fs = require('fs').promises;
const path = require('path');
const express = require('express');
const ExcelJS = require('exceljs');
const swaggerUi = require('swagger-ui-express');
const swaggerJsdoc = require('swagger-jsdoc');

const app = express();



app.use(express.json());


const options = {
    definition: {
        openapi: '3.0.0',
        info: {
            title: 'File Processing API',
            version: '1.0.0',
            description: 'API to process files and return an Excel file.',
        },
    },
    apis: ['./index.js'], 
};

const swaggerSpec = swaggerJsdoc(options);
app.use('/api-docs', swaggerUi.serve, swaggerUi.setup(swaggerSpec));

/**
 * @swagger
 * /files:
 *   post:
 *     summary: Process files in the given path and return an Excel file.
 *     requestBody:
 *       required: true
 *       content:
 *         application/json:
 *           schema:
 *             type: object
 *             properties:
 *               mainFolderPath:
 *                 type: string
 *                 description: The path to the main folder.
 *                 example: "C:/path/to/your/main/folder"
 *     responses:
 *       200:
 *         description: Excel file generated successfully.
 *         content:
 *           application/vnd.openxmlformats-officedocument.spreadsheetml.sheet:
 *             schema:
 *               type: string
 *               format: binary
 *       400:
 *         description: Invalid input provided.
 *       500:
 *         description: Internal server error.
 */

app.post('/files', async (req, res) => {
    try {
        const { mainFolderPath } = req.body;

        if (!mainFolderPath) {
            return res.status(400).json({ error: 'mainFolderPath is required in the request body.' });
        }

        const allResults = await processMainFolder(mainFolderPath);


        const workbook = await createExcelWorkbook(allResults);


        const buffer = await workbook.xlsx.writeBuffer();


        res.setHeader('Content-Disposition', 'attachment; filename=results.xlsx');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');


        res.send(buffer);
    } catch (error) {
        console.error('Error:', error);
        res.status(500).json({ error: 'An error occurred while processing the request.' });
    }
});

app.listen(port, () => {
    console.log(`Server is running on port ${port}`);
});


async function processMainFolder(mainFolderPath) {
    let results = [];

    try {
        const subfolders = await fs.readdir(mainFolderPath);

        for (const subfolder of subfolders) {
            const subfolderPath = path.join(mainFolderPath, subfolder);
            const stats = await fs.stat(subfolderPath);

            if (stats.isDirectory()) {
                const subfolderResults = await processSubfolder(subfolderPath);
                results = results.concat(subfolderResults);
            }
        }
    } catch (err) {
        console.error('Error processing main folder:', err);
        throw err;
    }


    let organizedResults = {};

    for (const result of results) {
        const modDate = result.modificationDate;
        if (!organizedResults[modDate]) {
            organizedResults[modDate] = [];
        }
        organizedResults[modDate].push(result);
    }

    return organizedResults;
}


async function processSubfolder(subfolderPath) {
    let results = [];

    try {
        const files = await fs.readdir(subfolderPath);

        let fileGroups = {};

        for (const file of files) {
            const filePath = path.join(subfolderPath, file);
            const stats = await fs.stat(filePath);

            if (stats.isFile()) {
                const match = file.match(/^(\d+)_(\d{2})(\d{2})(\d{4})_\d+_\d+\.dat$/);
                if (match) {
                    const magazineCode = match[1];
                    const dayFromFilename = match[2];
                    const monthFromFilename = match[3];
                    const yearFromFilename = match[4];
                    const dateFromFilename = `${dayFromFilename}/${monthFromFilename}/${yearFromFilename}`;

                    const modDate = stats.mtime;
                    const dayMod = modDate.getDate().toString().padStart(2, '0');
                    const monthMod = (modDate.getMonth() + 1).toString().padStart(2, '0');
                    const yearMod = modDate.getFullYear();
                    const modificationDate = `${yearMod}-${monthMod}-${dayMod}`; 

                    const key = `${modificationDate}_${magazineCode}_${dayFromFilename}${monthFromFilename}${yearFromFilename}`;

                    if (!fileGroups[key]) {
                        fileGroups[key] = [];
                    }

                    fileGroups[key].push({
                        filePath,
                        fileName: file,
                        magazineCode,
                        dateFromFilename,
                        modificationDate,
                        mtime: stats.mtime
                    });
                }
            }
        }


        for (let key in fileGroups) {
            const groupFiles = fileGroups[key];

            groupFiles.sort((a, b) => b.mtime - a.mtime);


            const latestFile = groupFiles[0];


            const lastLine = await getLastLine(latestFile.filePath);

            results.push({
                fileName: latestFile.fileName,
                magazineCode: latestFile.magazineCode,
                dateFromFilename: latestFile.dateFromFilename,
                modificationDate: latestFile.modificationDate,
                InValue: lastLine.InValue,
                outValue: lastLine.outValue
            });
        }

    } catch (err) {
        console.error('Error processing subfolder:', err);
        throw err;
    }

    return results;
}


async function getLastLine(filePath) {
    try {
        const data = await fs.readFile(filePath, 'utf8');
        const lines = data.trim().split('\n');
        const lastLine = lines[lines.length - 1].split('|');
        const InValue = lastLine[lastLine.length - 4];
        const outValue = lastLine[lastLine.length - 3];
        return {
            InValue,
            outValue
        };

    } catch (err) {
        console.error('Error reading file:', err);
        return {
            InValue: '',
            outValue: ''
        };
    }
}


async function createExcelWorkbook(data) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Results');

    worksheet.columns = [
        { header: 'Modification Date', key: 'modificationDate', width: 15 },
        { header: 'File Name', key: 'fileName', width: 30 },
        { header: 'Magazine Code', key: 'magazineCode', width: 15 },
        { header: 'Date From Filename', key: 'dateFromFilename', width: 15 },
        { header: 'In Value', key: 'InValue', width: 15 },
        { header: 'Out Value', key: 'outValue', width: 15 }
    ];


    for (const modDate in data) {
        const results = data[modDate];
        for (const result of results) {
            worksheet.addRow({
                modificationDate: result.modificationDate,
                fileName: result.fileName,
                magazineCode: result.magazineCode,
                dateFromFilename: result.dateFromFilename,
                InValue: result.InValue,
                outValue: result.outValue
            });
        }
    }

    return workbook;
}
