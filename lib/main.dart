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
  String? _currentOpeningFileName; // Lưu tên file thay vì full path để tránh lỗi permission

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
      _showSnackBar("Đã cài đặt đường dẫn lưu mặc định.");
    }
  }

  void _addNewRow() {
    setState(() => _controllers.add(List.generate(4, (_) => TextEditingController())));
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

      // ƯU TIÊN: Dùng saveFile của FilePicker để ghi đè hợp lệ theo chuẩn Android
      String? fileNameToSave;
      
      if (_currentOpeningFileName != null) {
        fileNameToSave = _currentOpeningFileName;
      } else {
        fileNameToSave = await _showFileNameDialog();
      }

      if (fileNameToSave == null || fileNameToSave.isEmpty) return;
      if (!fileNameToSave.endsWith('.xlsx')) fileNameToSave += '.xlsx';

      // Gọi lệnh lưu của hệ thống (Đây là cách chắc chắn nhất để ghi đè thành công)
      String? resultPath = await FilePicker.platform.saveFile(
        dialogTitle: 'Đang lưu file...',
        fileName: fileNameToSave,
        type: FileType.custom,
        allowedExtensions: ['xlsx'],
        bytes: bytes,
      );

      if (resultPath != null) {
        setState(() => _currentOpeningFileName = fileNameToSave);
        _showSnackBar("Lưu thành công!");
      }
    } catch (e) {
      _showSnackBar("Lỗi: $e");
    }
  }

  Future<String?> _showFileNameDialog() async {
    TextEditingController _nameCtrl = TextEditingController();
    return showDialog<String>(
      context: context,
      builder: (context) => AlertDialog(
        title: const Text("Tên file mới"),
        content: TextField(controller: _nameCtrl, decoration: const InputDecoration(hintText: "Nhập tên file..."), autofocus: true),
        actions: [
          TextButton(onPressed: () => Navigator.pop(context), child: const Text("Hủy")),
          ElevatedButton(onPressed: () => Navigator.pop(context, _nameCtrl.text), child: const Text("Lưu")),
        ],
      ),
    );
  }

  Future<void> _importExcel() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      type: FileType.custom, 
      allowedExtensions: ['xlsx'],
      initialDirectory: _defaultPath,
      withData: true,
    );
    
    if (result != null) {
      setState(() => _currentOpeningFileName = result.files.single.name);
      
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
          _showSnackBar("Đã mở: ${result.files.single.name}");
          break;
        }
      }
    }
  }

  void _showSnackBar(String message) {
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(message), duration: const Duration(seconds: 2)));
  }

  @override
  Widget build(BuildContext context) {
    bool isKeyboardVisible = MediaQuery.of(context).viewInsets.bottom != 0;

    return Scaffold(
      backgroundColor: Colors.white,
      resizeToAvoidBottomInset: true, 
      appBar: AppBar(
        title: const Text('Edit Excel', style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold)),
        centerTitle: true,
        flexibleSpace: Container(decoration: const BoxDecoration(gradient: LinearGradient(colors: [Colors.blue, Colors.indigo]))),
        actions: [
          IconButton(
            icon: const Icon(Icons.note_add, color: Colors.white),
            onPressed: () => setState(() {
              _controllers = [List.generate(4, (_) => TextEditingController())];
              _currentOpeningFileName = null;
              _showSnackBar("Đã tạo trang mới");
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
            width: double.infinity,
            color: _currentOpeningFileName == null ? Colors.orange[50] : Colors.green[50],
            padding: const EdgeInsets.symmetric(vertical: 6),
            child: Text(
              _currentOpeningFileName == null ? "🆕 Đang tạo file mới" : "📂 Ghi đè: $_currentOpeningFileName",
              textAlign: TextAlign.center, style: const TextStyle(fontSize: 12, fontWeight: FontWeight.bold)
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
                    children: row.map((c) => Padding(padding: const EdgeInsets.symmetric(horizontal: 4), child: TextField(controller: c, style: const TextStyle(fontSize: 14), decoration: const InputDecoration(border: InputBorder.none, hintText: "...")))).toList(),
                  )),
                ],
              ),
            ),
          ),
        ],
      ),
      floatingActionButton: isKeyboardVisible 
        ? null 
        : FloatingActionButton.extended(
            onPressed: _addNewRow,
            backgroundColor: Colors.indigo,
            label: const Text("Thêm dòng", style: TextStyle(color: Colors.white)),
            icon: const Icon(Icons.add, color: Colors.white)
          ),
    );
  }
}
