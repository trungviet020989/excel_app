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
  String? _currentOpeningFilePath;

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
      await [Permission.storage, Permission.manageExternalStorage].request();
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

      // XỬ LÝ GHI ĐÈ FILE CŨ
      if (_currentOpeningFilePath != null) {
        final file = File(_currentOpeningFilePath!);
        if (await file.exists()) {
          await file.writeAsBytes(fileBytes, mode: FileMode.write, flush: true); // Thêm flush: true để chắc chắn dữ liệu được lưu
          _showSnackBar("Đã ghi đè thành công!");
          return;
        }
      }

      // LƯU FILE MỚI
      String? customFileName = await _showFileNameDialog();
      if (customFileName == null || customFileName.isEmpty) return;
      String finalFileName = customFileName.endsWith('.xlsx') ? customFileName : "$customFileName.xlsx";

      if (_defaultPath != null) {
        final file = File("$_defaultPath/$finalFileName");
        await file.writeAsBytes(fileBytes, flush: true);
        setState(() => _currentOpeningFilePath = file.path);
        _showSnackBar("Đã lưu mới: $finalFileName");
      } else {
        String? selectedFile = await FilePicker.platform.saveFile(
          dialogTitle: 'Chọn nơi lưu', fileName: finalFileName,
          type: FileType.custom, allowedExtensions: ['xlsx'],
          bytes: Uint8List.fromList(fileBytes),
        );
        if (selectedFile != null) _showSnackBar("Lưu thành công!");
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
  );
  if (result != null) {
    String? filePath = result.files.first.path;
    if (filePath == null) {
      _showSnackBar("Không thể lấy đường dẫn file.");
      return;
    }

    // Lưu đúng đường dẫn gốc để sau này ghi đè
    setState(() => _currentOpeningFilePath = filePath);

    // Đọc dữ liệu từ file gốc
    Uint8List bytes = await File(filePath).readAsBytes();

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
      _showSnackBar("Đã mở: ${result.files.first.name}");
      break;
    }
  }
}


  void _showSnackBar(String message) {
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(message), duration: const Duration(seconds: 2)));
  }

  @override
  Widget build(BuildContext context) {
    // Kiểm tra xem bàn phím có đang hiện hay không
    bool isKeyboardVisible = MediaQuery.of(context).viewInsets.bottom != 0;

    return Scaffold(
      backgroundColor: Colors.white,
      // Khi bàn phím hiện, giao diện sẽ tự đẩy lên để không che khuất ô nhập
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
              _currentOpeningFilePath = null;
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
            color: _currentOpeningFilePath == null ? Colors.orange[50] : Colors.green[50],
            padding: const EdgeInsets.symmetric(vertical: 6),
            child: Text(
              _currentOpeningFilePath == null ? "🆕 Đang tạo file mới" : "📂 Ghi đè: ${_currentOpeningFilePath!.split('/').last}",
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
      // CHỈ HIỆN NÚT THÊM DÒNG KHI BÀN PHÍM ĐÃ ĐÓNG
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
