import express, { Request, Response } from 'express';
import ExcelJS from 'exceljs';
import * as fs from 'fs';
import * as path from 'path';
import dotenv from 'dotenv';

dotenv.config();


import { GoogleGenerativeAI } from '@google/generative-ai';
import { GoogleAIFileManager } from '@google/generative-ai/server';

const app = express();
const port = process.env.PORT || 3000;
app.use(express.json({ limit: '10mb' }));


const apiKey = process.env.GEMINI_API_KEY || 'YOUR_GEMINI_API_KEY';
const genAI = new GoogleGenerativeAI(apiKey);
const fileManager = new GoogleAIFileManager(apiKey);
const model = genAI.getGenerativeModel({
  model: "gemini-2.0-flash-thinking-exp-01-21",
});

async function uploadToGemini(filePath: string, mimeType: string) {
  const uploadResult = await fileManager.uploadFile(filePath, {
    mimeType,
    displayName: path.basename(filePath),
  });
  return uploadResult.file;
}

function saveExcelLocally(workbook: ExcelJS.Workbook, filename: string) {
  const localDir = path.join(__dirname, 'downloads');
  if (!fs.existsSync(localDir)) fs.mkdirSync(localDir, { recursive: true });
  const localPath = path.join(localDir, filename);
  return workbook.xlsx.writeFile(localPath);
}

app.post('/extract_data', async (req: Request, res: Response): Promise<any> => {
  try {
    const { image } = req.body;
    if (!image) return res.status(400).json({ error: 'Image data required' });

    
    const base64Data = image.split(',')[1] || image;
    const imageBuffer = Buffer.from(base64Data, 'base64');
    const tempFilePath = path.join(__dirname, 'temp.png');
    fs.writeFileSync(tempFilePath, imageBuffer);

   
    const uploadedFile = await uploadToGemini(tempFilePath, "image/png");
    
    
    const chatSession = model.startChat({
      generationConfig: {
        temperature: 0.7,
        topP: 0.95,
        topK: 64,
        maxOutputTokens: 65536,
        responseMimeType: "text/plain",
      },
      history: [{
        role: "user",
        parts: [
          { fileData: { mimeType: uploadedFile.mimeType, fileUri: uploadedFile.uri } },
          { text: `You are an expert at extracting tabular data from images. Analyze this table and follow these rules:

1. Column Structure:
   - The first 2 columns are ALWAYS: "Roll No" (a numeric or alphanumeric roll number) and "Name" (the full name).
   - The remaining columns represent attendance data, which are provided with two header rows:
       - The first header row contains the "Prac No" values.
       - The second header row contains handwritten date values in day/month format (e.g., "04/01").
   - Do NOT combine these header rows. Instead, output them as two separate header rows:
       - The first header row should list "Roll No", "Name", followed by the Prac No values (e.g., "Prac 1", "Prac 2", etc.).
       - The second header row should have empty values for the first two columns and then the corresponding date values (e.g., "04/01", "05/01", etc.).
   - Preserve the original column order from the image.

2. Data Extraction:
   - For the attendance cells, use "P" ONLY if it is a clear.
   - Treat any faint marks, shadows, or doubtful cells as null.
   - Preserve empty cells as null.
   - Please Scan Each column very sharply and mark "P" in the bracket only if it is visible.
   - There might be blank rows as well. Leave them blank only no need to add unnecessary 'P' there.
   
   

3. Validation:
   - Ensure that the final JSON includes both header rows and that all data rows have the same number of columns as defined by the headers.
   - Reject any non-tabular data or annotations.

Return the extracted table data as a JSON object with EXACT structure, using two keys: "headers" and "data". For example:

{
  "headers": [
    ["Roll No", "Name", "Prac 1", "Prac 2", "Prac 3"],
    ["", "", "04/01", "05/01", "06/01"]
  ],
  "data": [
    {
      "Roll No": "01",
      "Name": "John Doe",
      "Prac 1": "P",
      "Prac 2": null,
      "Prac 3": "P"
    },
    {
      "Roll No": "02",
      "Name": "Jane Smith",
      "Prac 1": null,
      "Prac 2": "P",
      "Prac 3": "P"
    }
  ]
}

          ` }
        ]
      }],
    });

   
    const result = await chatSession.sendMessage("");
    const rawJSON = result.response.text().replace(/```json|```/g, '').trim();
    const jsonData = JSON.parse(rawJSON);

 
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Data');

    
    if (jsonData.headers && jsonData.data) {
    
      jsonData.headers.forEach((headerRow: string[]) => {
        worksheet.addRow(headerRow);
      });

      
      const headerKeys = jsonData.headers[0];

      
      jsonData.data.forEach((row: any) => {
        const rowData = headerKeys.map((key: string) => row[key] ?? null);
        worksheet.addRow(rowData);
      });
    } else if (Array.isArray(jsonData)) {
      
      const columns = new Set<string>();
      jsonData.forEach((row: any) => Object.keys(row).forEach(k => columns.add(k)));
      const headers = Array.from(columns);
      worksheet.addRow(headers);
      jsonData.forEach((row: any) => {
        const rowData = headers.map(header => row[header] ?? null);
        worksheet.addRow(rowData);
      });
    } else {
      throw new Error("Invalid JSON format received from Gemini.");
    }

    const headerRow1 = worksheet.getRow(1);
    const totalColumns = headerRow1.cellCount; 
    const attendanceStart = 3; 

    let lastLectureCol = totalColumns;
    for (let col = totalColumns; col >= attendanceStart; col--) {
      let foundP = false;
      worksheet.eachRow((row, rowNumber) => {
        
        if (rowNumber > 2) {
          const cellValue = row.getCell(col).value;
          if (cellValue === "P") {
            foundP = true;
          }
        }
      });
      if (foundP) {
        lastLectureCol = col;
        break;
      }
    }
    const lectureCount = lastLectureCol - attendanceStart + 1;

    
    headerRow1.getCell(totalColumns + 1).value = "Attendance %";
    const headerRow2 = worksheet.getRow(2);
    headerRow2.getCell(totalColumns + 1).value = ""; 

    
    worksheet.eachRow((row, rowNumber) => {
      
      if (rowNumber > 2) {
        let presentCount = 0;
        for (let col = attendanceStart; col <= lastLectureCol; col++) {
          const cellValue = row.getCell(col).value;
          if (cellValue === "P") {
            presentCount++;
          }
        }
       
        const attendancePercent = lectureCount > 0 
          ? (presentCount / lectureCount) * 100 
          : 0;
        
        row.getCell(totalColumns + 1).value = `${attendancePercent.toFixed(2)}%`;
      }
    });

   
    const timestamp = new Date().toISOString()
      .replace(/[:.-]/g, '')
      .replace('T', '_')
      .slice(0, 15);
    const filename = `data_${timestamp}.xlsx`;

    await saveExcelLocally(workbook, filename);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    await workbook.xlsx.write(res);

    // Cleanup
    fs.unlinkSync(tempFilePath);
  } catch (error: any) {
    console.error("Processing error:", error);
    res.status(500).json({ error: error.message });
  }
});


app.listen(port, () => console.log(`Server running on port ${port}`));