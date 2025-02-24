import 'dart:convert';
import 'dart:io';
import 'package:flutter/material.dart';
import 'package:http/http.dart' as http;
import 'package:image_picker/image_picker.dart';
import 'package:path_provider/path_provider.dart';
import 'package:file_saver/file_saver.dart';

class HomePage extends StatefulWidget {
  const HomePage({super.key});

  @override
  _HomePageState createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {
  File? _image;
  bool _isLoading = false;

  Future<void> _pickImage() async {
    final pickedFile = await ImagePicker().pickImage(source: ImageSource.gallery);

    if (pickedFile != null) {
      setState(() {
        _image = File(pickedFile.path);
      });
    }
  }

  Future<void> _processImageAndDownloadExcel() async {
  if (_image == null) {
    ScaffoldMessenger.of(context).showSnackBar(
      const SnackBar(content: Text("Please select an image first!")),
    );
    return;
  }

  setState(() {
    _isLoading = true;
  });

  try {
    final url = Uri.parse("http://10.0.2.2:3000/extract_data"); // Replace with your backend URL
    List<int> imageBytes = await _image!.readAsBytes();
    String base64Image = base64Encode(imageBytes);

    final response = await http.post(
      url,
      headers: {'Content-Type': 'application/json'},
      body: jsonEncode({'image': base64Image}),
    );

    if (response.statusCode == 200) {
      final bytes = response.bodyBytes;

      
      await FileSaver.instance.saveFile(
        name: "data.xlsx",
        bytes: bytes,
        ext: "xlsx",
        mimeType: MimeType.other, // Corrected
      );

      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text("Excel file downloaded successfully!")),
      );
    } else {
      throw Exception('Failed to process image and download Excel');
    }
  } catch (e) {
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(content: Text("Error: ${e.toString()}")),
    );
  } finally {
    setState(() {
      _isLoading = false;
    });
  }
}


  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('Excel Generator')),
      body: Center(
        child: Padding(
          padding: const EdgeInsets.all(16.0),
          child: Column(
            mainAxisAlignment: MainAxisAlignment.center,
            children: [
              _image != null
                  ? Image.file(_image!, height: 200)
                  : const Text("No image selected", style: TextStyle(fontSize: 16)),
              const SizedBox(height: 20),
              ElevatedButton(
                onPressed: _pickImage,
                child: const Text("Pick Image"),
              ),
              const SizedBox(height: 10),
              ElevatedButton(
                onPressed: _isLoading ? null : _processImageAndDownloadExcel,
                child: _isLoading
                    ? const CircularProgressIndicator(color: Colors.white)
                    : const Text("Process & Download Excel"),
              ),
            ],
          ),
        ),
      ),
    );
  }
}
