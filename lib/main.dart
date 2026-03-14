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

  Future<void> _loadDefaultPath() async {
    final prefs = await SharedPreferences.getInstance();
    setState(() {
      _defaultPath = prefs.getString('default_path');
    });
  }

  // SỬA NÚT CÀI ĐẶT: Đảm bảo bảng chọn thư mục hiện lên
  Future<void> _settingsPath() async {
    // Xin quyền truy cập trước khi mở thư mục
    await [Permission.storage, Permission.manageExternalStorage].request();
    
    String? selectedDirectory = await FilePicker.platform.getDirectoryPath();
    if (selectedDirectory != null) {
      final prefs = await SharedPreferences.getInstance();
      await prefs.setString('default_path', selectedDirectory);
      setState(() => _defaultPath = selectedDirectory);
      _showSnackBar("Đã cài đặt đường dẫn: $selectedDirectory");
    }
  }

  void _addNewRow() {
    setState(() {
      _controllers.add(List.generate(4, (_) => TextEditingController()));
    });
  }

  // SỬA LỖI LƯU FILE: Xử lý lỗi "Bytes are required"
  Future<void> _exportExcel() async {
    try {
      await [Permission.storage, Permission.manageExternalStorage].request();

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

      // SỬA TẠI ĐÂY: Đảm bảo lấy được danh sách Bytes
      final List<int>? fileBytes = excel.save();
      if (fileBytes == null) throw "Không thể tạo dữ liệu file Excel";

      String fileName = "DuLieu_${DateTime.now().millisecondsSinceEpoch}.xlsx";

      if (_defaultPath != null) {
        final file = File("$_defaultPath/$fileName");
        await file.writeAsBytes(fileBytes);
        _showSnackBar("Đã lưu thành công tại: $_defaultPath");
      } else {
        String? selectedFile = await FilePicker.platform.saveFile(
          dialogTitle: 'Chọn nơi lưu file',
          fileName: fileName,
          type: FileType.custom,
          allowedExtensions: ['xlsx'],
          bytes: Uint8List.fromList(fileBytes), // Truyền bytes vào đây để tránh lỗi trên Android
        );
        if (selectedFile != null) _showSnackBar("Đã lưu thành công!");
      }
    } catch (e) {
      _showSnackBar("Lỗi: $e");
    }
  }

  Future<void> _importExcel() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
      initialDirectory: _defaultPath,
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
        title: const Text('Edit Excel', style: TextStyle(color: Colors.white)),
        centerTitle: true,
        flexibleSpace: Container(decoration: const BoxDecoration(gradient: LinearGradient(colors: [Colors.blue, Colors.indigo]))),
        actions: [
          IconButton(icon: const Icon(Icons.settings, color: Colors.white), onPressed: _settingsPath),
          IconButton(icon: const Icon(Icons.file_open, color: Colors.white), onPressed: _importExcel),
          IconButton(icon: const Icon(Icons.save, color: Colors.white), onPressed: _exportExcel),
        ],
      ),
      body: Column(
        children: [
          if (_defaultPath != null)
            Container(width: double.infinity, color: Colors.green[50], padding: const EdgeInsets.all(5),
              child: Text("📂 Thư mục lưu: $_defaultPath", textAlign: TextAlign.center, style: const TextStyle(fontSize: 10))),
          Expanded(
            child: SingleChildScrollView(
              padding: const EdgeInsets.all(10),
              child: Card(
                child: Table(
                  border: TableBorder.all(color: Colors.grey.shade300),
                  children: [
                    TableRow(
                      decoration: const BoxDecoration(color: Colors.indigo),
                      children: ['Tên SP', 'Giá Bán', 'Giá Nhập', 'SL'].map((t) => Padding(padding: const EdgeInsets.all(8), child: Text(t, style: const TextStyle(color: Colors.white, fontWeight: FontWeight.bold)))).toList(),
                    ),
                    ..._controllers.map((row) => TableRow(
                      children: row.map((c) => Padding(padding: const EdgeInsets.symmetric(horizontal: 5), child: TextField(controller: c, decoration: const InputDecoration(border: InputBorder.none)))).toList(),
                    )),
                  ],
                ),
              ),
            ),
          ),
        ],
      ),
      floatingActionButton: FloatingActionButton(onPressed: _addNewRow, backgroundColor: Colors.indigo, child: const Icon(Icons.add, color: Colors.white)),
    );
  }
}
