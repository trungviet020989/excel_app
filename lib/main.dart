import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart' as ex;
import 'dart:typed_data';
import 'dart:io';
import 'package:shared_preferences/shared_preferences.dart';
import 'package:permission_handler/permission_handler.dart';

void main() => runApp(const MaterialApp(home: ExcelApp(), debugShowCheckedModeBanner: false));

class ExcelApp extends StatefulWidget {
  const ExcelApp({super.key});
  @override
  State<ExcelApp> createState() => _ExcelAppState();
}

class _ExcelAppState extends State<ExcelApp> {
  List<List<TextEditingController>> _controllers = [];
  String? _defaultPath;

  @override
  void initState() {
    super.initState();
    _addNewRow();
    _loadDefaultPath();
  }

  // Tải đường dẫn đã cài đặt từ bộ nhớ máy
  Future<void> _loadDefaultPath() async {
    final prefs = await SharedPreferences.getInstance();
    setState(() {
      _defaultPath = prefs.getString('default_path');
    });
  }

  // Nút Cài Đặt: Chọn và lưu đường dẫn mặc định
  Future<void> _settingsPath() async {
    if (await Permission.storage.request().isGranted || await Permission.manageExternalStorage.request().isGranted) {
      String? selectedDirectory = await FilePicker.platform.getDirectoryPath();
      if (selectedDirectory != null) {
        final prefs = await SharedPreferences.getInstance();
        await prefs.setString('default_path', selectedDirectory);
        setState(() => _defaultPath = selectedDirectory);
        _showSnackBar("Đã cài đặt đường dẫn: $selectedDirectory");
      }
    }
  }

  void _addNewRow() {
    setState(() {
      _controllers.add(List.generate(4, (_) => TextEditingController()));
    });
  }

  // CHỨC NĂNG LƯU FILE
  Future<void> _exportExcel() async {
    try {
      // Yêu cầu quyền truy cập bộ nhớ
      var status = await Permission.storage.request();
      if (!status.isGranted) {
        await Permission.manageExternalStorage.request();
      }

      var excel = ex.Excel.createExcel();
      ex.Sheet sheetObject = excel['Sheet1'];

      sheetObject.appendRow([
        ex.TextCellValue('Tên Sản Phẩm'),
        ex.TextCellValue('Giá Bán'),
        ex.TextCellValue('Giá Nhập'),
        ex.TextCellValue('Số Lượng'),
      ]);

      for (var row in _controllers) {
        sheetObject.appendRow([
          ex.TextCellValue(row[0].text),
          ex.TextCellValue(row[1].text),
          ex.TextCellValue(row[2].text),
          ex.TextCellValue(row[3].text),
        ]);
      }

      var bytes = excel.encode();
      String fileName = "DuLieu_${DateTime.now().millisecondsSinceEpoch}.xlsx";
      String? fullPath;

      if (_defaultPath != null) {
        // Lưu vào đường dẫn đã cài đặt
        fullPath = "$_defaultPath/$fileName";
        final file = File(fullPath);
        await file.writeAsBytes(bytes!);
        _showSnackBar("Đã lưu thành công tại: $fullPath");
      } else {
        // Nếu chưa cài đường dẫn, hiện bảng chọn
        String? selectedFile = await FilePicker.platform.saveFile(
          dialogTitle: 'Chọn nơi lưu file',
          fileName: fileName,
          type: FileType.custom,
          allowedExtensions: ['xlsx'],
        );
        if (selectedFile != null) {
          final file = File(selectedFile);
          await file.writeAsBytes(bytes!);
          _showSnackBar("Đã lưu thành công!");
        }
      }
    } catch (e) {
      _showSnackBar("Lỗi lưu file: $e");
    }
  }

  // CHỨC NĂNG MỞ FILE
  Future<void> _importExcel() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
      initialDirectory: _defaultPath, // Mở ngay tại đường dẫn cài đặt
      withData: true,
    );
    
    if (result != null) {
      Uint8List? bytes = result.files.first.bytes;
      if (bytes != null) {
        var excel = ex.Excel.decodeBytes(bytes);
        for (var table in excel.tables.keys) {
          setState(() {
            _controllers.clear();
            var tableData = excel.tables[table]!;
            for (int i = 1; i < tableData.rows.length; i++) {
              var rowData = tableData.rows[i];
              _controllers.add([
                TextEditingController(text: rowData[0]?.value?.toString() ?? ""),
                TextEditingController(text: rowData[1]?.value?.toString() ?? ""),
                TextEditingController(text: rowData[2]?.value?.toString() ?? ""),
                TextEditingController(text: rowData[3]?.value?.toString() ?? ""),
              ]);
            }
          });
          _showSnackBar("Đã nhập dữ liệu thành công!");
          break; 
        }
      }
    }
  }

  void _showSnackBar(String message) {
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(message)));
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: Colors.grey[100],
      appBar: AppBar(
        title: const Text('Edit Excel', style: TextStyle(fontWeight: FontWeight.bold)),
        centerTitle: true,
        elevation: 5,
        flexibleSpace: Container(
          decoration: const BoxDecoration(
            gradient: LinearGradient(colors: [Colors.blueAccent, Colors.indigo])
          ),
        ),
        actions: [
          IconButton(icon: const Icon(Icons.settings, color: Colors.white), onPressed: _settingsPath, tooltip: "Cài Đặt"),
          IconButton(icon: const Icon(Icons.file_open, color: Colors.white), onPressed: _importExcel),
          IconButton(icon: const Icon(Icons.save, color: Colors.white), onPressed: _exportExcel),
        ],
      ),
      body: Column(
        children: [
          if (_defaultPath != null)
            Container(
              padding: const EdgeInsets.all(8),
              color: Colors.yellow[100],
              width: double.infinity,
              child: Text("📂 Thư mục mặc định: ${_defaultPath!.split('/').last}", 
                textAlign: TextAlign.center, style: const TextStyle(fontSize: 12, color: Colors.brown)),
            ),
          Expanded(
            child: SingleChildScrollView(
              padding: const EdgeInsets.all(12),
              child: Card(
                elevation: 4,
                shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
                child: Table(
                  border: TableBorder.symmetric(inside: BorderSide(color: Colors.grey.shade300)),
                  columnWidths: const {
                    0: FlexColumnWidth(2.5),
                    1: FlexColumnWidth(1.5),
                    2: FlexColumnWidth(1.5),
                    3: FlexColumnWidth(1.2),
                  },
                  children: [
                    TableRow(
                      decoration: BoxDecoration(
                        color: Colors.indigo[400],
                        borderRadius: const BorderRadius.vertical(top: Radius.circular(12)),
                      ),
                      children: ['Tên SP', 'Giá Bán', 'Giá Nhập', 'SL'].map((text) => 
                        Padding(
                          padding: const EdgeInsets.all(12),
                          child: Text(text, style: const TextStyle(color: Colors.white, fontWeight: FontWeight.bold, fontSize: 13))
                        )
                      ).toList(),
                    ),
                    ..._controllers.map((rowControllers) => TableRow(
                      children: rowControllers.map((ctrl) => Padding(
                        padding: const EdgeInsets.symmetric(horizontal: 8),
                        child: TextField(
                          controller: ctrl, 
                          style: const TextStyle(fontSize: 14),
                          decoration: const InputDecoration(border: InputBorder.none, hintText: '...'),
                        ),
                      )).toList(),
                    )),
                  ],
                ),
              ),
            ),
          ),
        ],
      ),
      floatingActionButton: FloatingActionButton.extended(
        onPressed: _addNewRow, 
        backgroundColor: Colors.indigo,
        label: const Text("Thêm dòng"),
        icon: const Icon(Icons.add),
      ),
    );
  }
}
