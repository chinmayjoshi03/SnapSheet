# 📊 Snapsheet - Attendance Sheet Image to SpreadSheet Converter

This project allows users to upload an image containing a table, extract the data using Google's Gemini AI, and download the extracted data as an Excel file. It consists of a Flutter frontend for image selection and file download, and a Node.js backend for processing the image and generating the Excel file.

## ✨ Features

- 📷 **Image Upload**: Select an image from your device.
- 🤖 **AI-Powered Data Extraction**: Use Google's Gemini AI to extract table data from the image.
- 👥 **Excel Download**: Download the extracted data as an Excel file.
- ⚠️ **Error Handling**: Robust error handling for invalid inputs and API failures.

## 🛠️ Tech Stack

### Frontend (Flutter)

- 📱 **Flutter**: For building the cross-platform mobile app.
- 🌐 **http**: For making API requests to the backend.
- 🖼️ **image_picker**: For selecting images from the device.
- 💾 **file_saver**: For saving the Excel file to the device.

### Backend (Node.js)

- ✨ **Express.js**: For handling HTTP requests.
- 🤓 **Google Generative AI**: For extracting table data from images.
- 📚 **ExcelJS**: For generating Excel files.
- 🔒 **dotenv**: For managing environment variables.
- 🌍 **cors**: For enabling cross-origin requests.

## ⚙️ Setup Instructions

### Prerequisites

- **Flutter SDK**: Install Flutter from [flutter.dev](https://flutter.dev/).
- **Node.js**: Install Node.js from [nodejs.org](https://nodejs.org/).
- **Google Gemini API Key**: Obtain an API key from [Google AI Studio](https://aistudio.google.com/).

### Backend Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/table-data-extractor.git
   cd table-data-extractor/backend
   ```
2. Install dependencies:
   ```bash
   npm install
   ```
3. Create a `.env` file in the backend directory and add your Gemini API key:
   ```env
   GEMINI_API_KEY=your_api_key_here
   PORT=3000
   ```
4. Start the backend server:
   ```bash
   npm start
   ```
   The backend will run at `http://localhost:3000`.

### Frontend Setup

1. Navigate to the frontend directory:
   ```bash
   cd ../frontend
   ```
2. Install dependencies:
   ```bash
   flutter pub get
   ```
3. Update the backend URL in `lib/home_page.dart`:
   ```dart
   final url = Uri.parse('http://localhost:3000/extract_data');
   ```
4. Run the Flutter app:
   ```bash
   flutter run
   ```

## 📰 API Endpoints

### **POST /extract_data**

- **Description**: Extracts table data from an image and returns an Excel file.
- **Request Body:**
  ```json
  {
    "image": "base64_encoded_image"
  }
  ```
- **Response:**
  - Success: Excel file (`application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`).
  - Error: JSON with error message.

## 💪 Contributing

1. Fork the repository.
2. Create a new branch:
   ```bash
   git checkout -b feature/your-feature-name
   ```
3. Commit your changes:
   ```bash
   git commit -m "Add your feature"
   ```
4. Push to the branch:
   ```bash
   git push origin feature/your-feature-name
   ```
5. Open a pull request.

## ❓ Support

For issues or questions, please open an issue on the GitHub repository.

Enjoy extracting table data with ease! 🚀

## 📧 Contact & Support
For questions, feel free to open an **issue** or contact me at **chinmayjoshi003@gmail.com** 📩

---