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
  String? _currentFileNameOnly;

  @override
  void initState() {
    super.initState();
    _addNewRow();
    _loadDefaultPath();
  }

  Future<void> _loadDefaultPath() async {
    final prefs = await SharedPreferences.getInstance();
    setState(() => _defaultPath = prefs.getString('default_path'));
  }

  Future<void> _settingsPath() async {
    await [Permission.storage, Permission.manageExternalStorage].request();
    String? selectedDirectory = await FilePicker.platform.getDirectoryPath();
    if (selectedDirectory != null) {
      final prefs = await SharedPreferences.getInstance();
      await prefs.setString('default_path', selectedDirectory);
      setState(() => _defaultPath = selectedDirectory);
      _showSnackBar("Đã cài đặt thư mục mặc định.");
    }
  }

  void _addNewRow() {
    setState(() => _controllers.add(List.generate(4, (_) => TextEditingController())));
  }

  String _suggestNextFileName() {
    if (_currentFileNameOnly == null) return "";
    if (_defaultPath == null) return _currentFileNameOnly!;
    
    String baseName = _currentFileNameOnly!;
    RegExp regExp = RegExp(r"^(.*?)(\d*)$");
    var match = regExp.firstMatch(baseName);
    String rootName = match?.group(1) ?? baseName;

    int counter = 1;
    String lastDigits = match?.group(2) ?? "";
    if (lastDigits.isNotEmpty) {
      counter = int.parse(lastDigits) + 1;
    }

    while (true) {
      String checkName = "$rootName$counter.xlsx";
      if (!File("$_defaultPath/$checkName").existsSync()) {
        return "$rootName$counter";
      }
      counter++;
    }
  }

  Future<void> _exportExcel() async {
    try {
      if (Platform.isAndroid) {
        await [Permission.storage, Permission.manageExternalStorage].request();
      }

      var excel = ex.Excel.createExcel();
      ex.Sheet sheetObject = excel['Sheet1'];

      sheetObject.appendRow([
        ex.TextCellValue('Tên Sản Phẩm'), ex.TextCellValue('Giá Bán'),
        ex.TextCellValue('Giá Nhập'), ex.TextCellValue('Số Lượng'),
      ]);

      for (var row in _controllers) {
        sheetObject.appendRow([
          ex.TextCellValue(row[0].text), ex.TextCellValue(row[1].text),
          ex.TextCellValue(row[2].text), ex.TextCellValue(row[3].text),
        ]);
      }

      final List<int>? fileBytes = excel.save();
      if (fileBytes == null) return;
      Uint8List bytes = Uint8List.fromList(fileBytes);

      String suggestion = _suggestNextFileName();
      String? customFileName = await _showFileNameDialog(suggestion);
      
      if (customFileName == null || customFileName.isEmpty) return;
      String finalFileName = customFileName.endsWith('.xlsx') ? customFileName : "$customFileName.xlsx";

      if (_defaultPath != null) {
        final file = File("$_defaultPath/$finalFileName");
        await file.writeAsBytes(bytes, flush: true);
        setState(() => _currentFileNameOnly = customFileName.replaceAll('.xlsx', ''));
        _showSnackBar("Đã lưu: $finalFileName");
      } else {
        String? selectedFile = await FilePicker.platform.saveFile(
          dialogTitle: 'Lưu file',
          fileName: finalFileName,
          type: FileType.custom,
          allowedExtensions: ['xlsx'],
          bytes: bytes,
        );
        if (selectedFile != null) {
          setState(() => _currentFileNameOnly = customFileName.replaceAll('.xlsx', ''));
          _showSnackBar("Lưu thành công!");
        }
      }
    } catch (e) {
      _showSnackBar("Lỗi khi lưu: $e");
    }
  }

  Future<String?> _showFileNameDialog(String initialName) async {
    TextEditingController _nameCtrl = TextEditingController(text: initialName);
    return showDialog<String>(
      context: context,
      builder: (context) => AlertDialog(
        title: Text(initialName.isEmpty ? "Lưu file mới" : "Lưu bản sao (Save As)"),
        content: TextField(
          controller: _nameCtrl, 
          decoration: const InputDecoration(hintText: "Nhập tên file...", suffixText: ".xlsx"),
          autofocus: true
        ),
        actions: [
          TextButton(onPressed: () => Navigator.pop(context), child: const Text("Hủy")),
          ElevatedButton(onPressed: () => Navigator.pop(context, _nameCtrl.text), child: const Text("Xác nhận")),
        ],
      ),
    );
  }

  Future<void> _importExcel() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      type: FileType.custom, 
      allowedExtensions: ['xlsx'],
      withData: true,
    );
    
    if (result != null) {
      String fileName = result.files.single.name;
      setState(() => _currentFileNameOnly = fileName.split('.').first);
      
      Uint8List? bytes = result.files.single.bytes;
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
          _showSnackBar("Đã mở: $fileName");
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
    bool isKeyboardVisible = MediaQuery.of(context).viewInsets.bottom != 0;

    return Scaffold(
      appBar: AppBar(
        title: const Text('Excel Manager', style: TextStyle(color: Colors.white, fontSize: 18)),
        flexibleSpace: Container(decoration: const BoxDecoration(gradient: LinearGradient(colors: [Colors.blue, Colors.indigo]))),
        actions: [
          IconButton(
            icon: const Icon(Icons.note_add, color: Colors.white),
            onPressed: () => setState(() {
              _controllers = [List.generate(4, (_) => TextEditingController())];
              _currentFileNameOnly = null;
              _showSnackBar("Trang mới");
            }),
          ),
          IconButton(icon: const Icon(Icons.settings, color: Colors.white), onPressed: _settingsPath),
          IconButton(icon: const Icon(Icons.file_open, color: Colors.white), onPressed: _importExcel),
          IconButton(icon: const Icon(Icons.save, color: Colors.white), onPressed: _exportExcel),
        ],
      ),
      body: Column(
        children: [
          Container(
            padding: const EdgeInsets.all(8),
            color: Colors.grey[100],
            child: Row(
              children: [
                Icon(Icons.edit_document, size: 16, color: _currentFileNameOnly == null ? Colors.orange : Colors.green),
                const SizedBox(width: 8),
                Text(_currentFileNameOnly == null ? "Tệp mới (Chưa lưu)" : "Đang chỉnh sửa: $_currentFileNameOnly.xlsx", 
                     style: const TextStyle(fontSize: 12, fontWeight: FontWeight.bold)),
              ],
            ),
          ),
          Expanded(
            child: SingleChildScrollView(
              padding: const EdgeInsets.all(8),
              child: Table(
                border: TableBorder.all(color: Colors.grey.shade300),
                children: [
                  TableRow(
                    decoration: const BoxDecoration(color: Colors.indigo),
                    children: ['Tên SP', 'Giá Bán', 'Giá Nhập', 'SL'].map((t) => Padding(padding: const EdgeInsets.all(10), child: Text(t, style: const TextStyle(color: Colors.white, fontWeight: FontWeight.bold, fontSize: 13)))).toList(),
                  ),
                  ..._controllers.map((row) => TableRow(
                    children: [
                      _buildTableCell(row[0], TextInputType.text),
                      _buildTableCell(row[1], TextInputType.number),
                      _buildTableCell(row[2], TextInputType.number),
                      _buildTableCell(row[3], TextInputType.number),
                    ],
                  )),
                ],
              ),
            ),
          ),
        ],
      ),
      // THAY ĐỔI TẠI ĐÂY: Nút thêm dòng gọn nhẹ chỉ có dấu +
      floatingActionButton: isKeyboardVisible 
        ? null 
        : FloatingActionButton(
            onPressed: _addNewRow,
            backgroundColor: Colors.indigo,
            child: const Icon(Icons.add, color: Colors.white),
          ),
    );
  }

  Widget _buildTableCell(TextEditingController controller, TextInputType keyboardType) {
    return Padding(
      padding: const EdgeInsets.symmetric(horizontal: 4),
      child: TextField(
        controller: controller,
        keyboardType: keyboardType,
        style: const TextStyle(fontSize: 14),
        decoration: const InputDecoration(
          border: InputBorder.none,
          hintText: "...",
        ),
      ),
    );
  }
}
