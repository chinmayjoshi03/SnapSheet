import express, { Request, Response } from 'express';
import ExcelJS from 'exceljs';
import * as fs from 'fs';
import * as path from 'path';
import dotenv from 'dotenv';

dotenv.config();

// Use a higher limit if needed
import * as base64 from 'base64-arraybuffer';

// Import Google Generative AI modules
import { GoogleGenerativeAI } from '@google/generative-ai';
import { GoogleAIFileManager } from '@google/generative-ai/server';

const app = express();
const port = process.env.PORT || 3000;
app.use(express.json({ limit: '10mb' })); // Adjust limit as needed

// Set up your Gemini API key (ensure it's set in your environment or replace with your key)
const apiKey = process.env.GEMINI_API_KEY || 'YOUR_GEMINI_API_KEY';

// Initialize the Generative AI and File Manager
const genAI = new GoogleGenerativeAI(apiKey);
const fileManager = new GoogleAIFileManager(apiKey);

/**
 * Uploads a local file to Gemini using the File Manager.
 */
async function uploadToGemini(filePath: string, mimeType: string) {
  const uploadResult = await fileManager.uploadFile(filePath, {
    mimeType,
    displayName: path.basename(filePath),
  });
  const file = uploadResult.file;
  console.log(`Uploaded file ${file.displayName} as: ${file.name}`);
  return file;
}

function saveExcelLocally(workbook: ExcelJS.Workbook, filename: string) {
    const localDir = path.join(__dirname, 'downloads');

    // Create directory if not exists
    if (!fs.existsSync(localDir)) {
        fs.mkdirSync(localDir, { recursive: true });
    }
    const localPath = path.join(localDir, filename);
    return workbook.xlsx.writeFile(localPath)
        .then(() => console.log(`File saved locally at: ${localPath}`))
        .catch((err) => console.error('Error saving file locally:', err));
}

// Create the generative model instance
const model = genAI.getGenerativeModel({
  model: "gemini-2.0-flash-thinking-exp-01-21",
});

app.post('/extract_data', async (req: Request, res: Response): Promise<any> => {
  try {
    const { image } = req.body;
    if (!image) {
      return res.status(400).json({ error: 'Image data is required' });
    }

    // Decode the base64 image and write it to a temporary file
    const base64Data = image.startsWith('data:')  ? image.split(',')[1]   : image;
    if (!base64Data) {
        return res.status(400).json({ error: 'Invalid image format' });
      }
  
    const imageBuffer = Buffer.from(base64Data, 'base64');
    if (!imageBuffer || imageBuffer.length === 0) {
        return res.status(400).json({ error: 'Empty image data' });
      }
      
    const tempFilePath = path.join(__dirname, 'temp.png');
    fs.writeFileSync(tempFilePath, imageBuffer);

    // Upload the file to Gemini
    const uploadedFile = await uploadToGemini(tempFilePath, "image/png");

    // Set up the generation configuration
    const generationConfig = {
      temperature: 0.7,
      topP: 0.95,
      topK: 64,
      maxOutputTokens: 65536,
      responseMimeType: "text/plain",
    };

    // Start a chat session with Gemini, including the file and the prompt in the history
    const chatSession = model.startChat({
      generationConfig,
      history: [
        {
          role: "user",
          parts: [
            {
              fileData: {
                mimeType: uploadedFile.mimeType,
                fileUri: uploadedFile.uri,
              },
            },
            {
              text:
                "You are an expert at extracting data from tables in images. Analyze the image and extract the information into a JSON format. The table contains information with the following structure:\n\n" +
                "Each row represents a person. The columns are 'BIB ID', 'Name', 'Col1', 'Col2', ..., 'Col10'. 'BIB ID' starts with 'BIB' followed by a number (e.g., 'BIB02'). 'Name' contains the person's full name. 'Col1' to 'Col10' contain either 'P' or are blank.\n\n" +
                "Return a JSON array where each object represents a row."
            }
          ],
        }
      ],
    });

    // Send the chat message (an empty string is used since the prompt is in the history)
    const result = await chatSession.sendMessage("");
    console.log("Gemini response:", result.response.text());

    const rawResponse = result.response.text();
    const cleanedResponse = rawResponse.replace(/```json/g, '').replace(/```/g, '').trim(); 

    const jsonData = JSON.parse(cleanedResponse);

    // Create an Excel workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Data');

    // Add header row
    worksheet.addRow([
      'BIB ID', 'Name', 'Col1', 'Col2', 'Col3', 'Col4',
      'Col5', 'Col6', 'Col7', 'Col8', 'Col9', 'Col10'
    ]);

    // Populate the worksheet with rows from the JSON data
    jsonData.forEach((row: any) => {
      worksheet.addRow([
        row['BIB ID'],
        row['Name'],
        row['Col1'],
        row['Col2'],
        row['Col3'],
        row['Col4'],
        row['Col5'],
        row['Col6'],
        row['Col7'],
        row['Col8'],
        row['Col9'],
        row['Col10'],
      ]);
    });

    await saveExcelLocally(workbook, "data.xlsx");
    
    // Set response headers for Excel file download
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=data.xlsx');

    // Write the workbook to the response stream
    await workbook.xlsx.write(res);
    res.end();
    
    // Clean up the temporary file
    fs.unlinkSync(tempFilePath);
  } catch (error: any) {
    console.error("Error processing image:", error);
    res.status(500).json({ error: error.message });
  }
});

app.listen(port, () => {
  console.log(`Server listening at http://localhost:${port}`);
});
