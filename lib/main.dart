import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart' as ex;
import 'dart:typed_data';
import 'dart:io';
import 'package:shared_preferences/shared_preferences.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:share_plus/share_plus.dart';

void main() => runApp(MaterialApp(
      home: const ExcelApp(),
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        primarySwatch: Colors.deepOrange,
        useMaterial3: true, // Sử dụng giao diện hiện đại hơn
      ),
    ));

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

  Future<void> _pickAndShareFile() async {
    if (_defaultPath == null) {
      _showSnackBar("Vui lòng cài đặt thư mục mặc định trước.");
      return;
    }
    final directory = Directory(_defaultPath!);
    if (!await directory.exists()) {
      _showSnackBar("Thư mục không tồn tại.");
      return;
    }
    List<FileSystemEntity> files = directory.listSync()
        .where((file) => file.path.endsWith('.xlsx'))
        .toList();

    if (files.isEmpty) {
      _showSnackBar("Không tìm thấy file Excel nào.");
      return;
    }

    showDialog(
      context: context,
      builder: (context) => AlertDialog(
        title: const Text("Gửi file sản phẩm", style: TextStyle(color: Colors.deepOrange, fontWeight: FontWeight.bold)),
        content: SizedBox(
          width: double.maxFinite,
          child: ListView.builder(
            shrinkWrap: true,
            itemCount: files.length,
            itemBuilder: (context, index) {
              String fileName = files[index].path.split('/').last;
              return Card( // Làm các mục trong danh sách nổi lên
                elevation: 2,
                child: ListTile(
                  leading: const Icon(Icons.description, color: Colors.orange),
                  title: Text(fileName, style: const TextStyle(fontSize: 14)),
                  onTap: () async {
                    Navigator.pop(context);
                    await Share.shareXFiles([XFile(files[index].path)]);
                  },
                ),
              );
            },
          ),
        ),
      ),
    );
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
    if (lastDigits.isNotEmpty) counter = int.parse(lastDigits) + 1;

    while (true) {
      String checkName = "$rootName$counter.xlsx";
      if (!File("$_defaultPath/$checkName").existsSync()) return "$rootName$counter";
      counter++;
    }
  }

  Future<void> _exportExcel() async {
    try {
      if (Platform.isAndroid) await [Permission.storage, Permission.manageExternalStorage].request();
      var excel = ex.Excel.createExcel();
      ex.Sheet sheetObject = excel['Sheet1'];
      sheetObject.appendRow([ex.TextCellValue('Tên SP'), ex.TextCellValue('Giá Bán'), ex.TextCellValue('Giá Nhập'), ex.TextCellValue('SL')]);
      for (var row in _controllers) {
        sheetObject.appendRow([ex.TextCellValue(row[0].text), ex.TextCellValue(row[1].text), ex.TextCellValue(row[2].text), ex.TextCellValue(row[3].text)]);
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
      }
    } catch (e) { _showSnackBar("Lỗi: $e"); }
  }

  Future<String?> _showFileNameDialog(String initialName) async {
    TextEditingController _nameCtrl = TextEditingController(text: initialName);
    return showDialog<String>(
      context: context,
      builder: (context) => AlertDialog(
        title: Text(initialName.isEmpty ? "Lưu file mới" : "Lưu bản sao"),
        content: TextField(controller: _nameCtrl, decoration: const InputDecoration(suffixText: ".xlsx"), autofocus: true),
        actions: [
          TextButton(onPressed: () => Navigator.pop(context), child: const Text("Hủy")),
          ElevatedButton(onPressed: () => Navigator.pop(context, _nameCtrl.text), child: const Text("Xác nhận")),
        ],
      ),
    );
  }

  Future<void> _importExcel() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles(type: FileType.custom, allowedExtensions: ['xlsx'], withData: true);
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
          break;
        }
      }
    }
  }

  void _showSnackBar(String message) {
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(message), backgroundColor: Colors.deepOrange));
  }

  @override
  Widget build(BuildContext context) {
    bool isKeyboardVisible = MediaQuery.of(context).viewInsets.bottom != 0;

    return Scaffold(
      backgroundColor: Colors.orange[50], // Màu nền app cam nhạt cực kỳ tươi
      appBar: AppBar(
        title: const Text('QUẢN LÝ KHO', style: TextStyle(color: Colors.white, fontSize: 18, fontWeight: FontWeight.bold, letterSpacing: 1.2)),
        centerTitle: true,
        flexibleSpace: Container(
          decoration: const BoxDecoration(
            gradient: LinearGradient(colors: [Colors.orange, Colors.deepOrange]) // Hiệu ứng Gradient cam
          )
        ),
        actions: [
          IconButton(icon: const Icon(Icons.share, color: Colors.white), onPressed: _pickAndShareFile),
          IconButton(icon: const Icon(Icons.settings, color: Colors.white), onPressed: _settingsPath),
          IconButton(icon: const Icon(Icons.file_open, color: Colors.white), onPressed: _importExcel),
          IconButton(icon: const Icon(Icons.save, color: Colors.white), onPressed: _exportExcel),
        ],
      ),
      body: Column(
        children: [
          // Thanh hiển thị tên file đẹp hơn
          Container(
            width: double.infinity,
            padding: const EdgeInsets.symmetric(vertical: 10, horizontal: 15),
            decoration: BoxDecoration(
              color: Colors.white,
              boxShadow: [BoxShadow(color: Colors.black12, blurRadius: 4, offset: const Offset(0, 2))]
            ),
            child: Row(
              children: [
                const Icon(Icons.folder_open, size: 18, color: Colors.orange),
                const SizedBox(width: 10),
                Expanded(
                  child: Text(
                    _currentFileNameOnly == null ? "Đang tạo file mới..." : "File hiện tại: $_currentFileNameOnly.xlsx",
                    style: TextStyle(color: Colors.orange[800], fontWeight: FontWeight.w600, fontSize: 13),
                  ),
                ),
              ],
            ),
          ),
          Expanded(
            child: SingleChildScrollView(
              padding: const EdgeInsets.all(10),
              child: Card( // Bọc bảng vào Card để tạo hiệu ứng nổi
                elevation: 3,
                shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(10)),
                child: ClipRRect(
                  borderRadius: BorderRadius.circular(10),
                  child: Table(
                    columnWidths: const {
                      0: FlexColumnWidth(2.5), // Tên SP rộng hơn chút
                      1: FlexColumnWidth(1.5),
                      2: FlexColumnWidth(1.5),
                      3: FlexColumnWidth(1),
                    },
                    border: TableBorder.all(color: Colors.orange.shade100, width: 0.5),
                    children: [
                      TableRow(
                        decoration: const BoxDecoration(color: Colors.orange),
                        children: ['Tên SP', 'Giá Bán', 'Giá Nhập', 'SL'].map((t) => Padding(
                          padding: const EdgeInsets.all(12), 
                          child: Text(t, style: const TextStyle(color: Colors.white, fontWeight: FontWeight.bold, fontSize: 12))
                        )).toList(),
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
            ),
          ),
        ],
      ),
      floatingActionButton: isKeyboardVisible ? null : FloatingActionButton(
        onPressed: _addNewRow,
        backgroundColor: Colors.deepOrange,
        elevation: 6,
        child: const Icon(Icons.add, color: Colors.white, size: 30),
      ),
    );
  }

  Widget _buildTableCell(TextEditingController controller, TextInputType keyboardType) {
    return Padding(
      padding: const EdgeInsets.symmetric(horizontal: 5),
      child: TextField(
        controller: controller,
        keyboardType: keyboardType,
        style: const TextStyle(fontSize: 14),
        textAlign: keyboardType == TextInputType.number ? TextAlign.center : TextAlign.left,
        decoration: const InputDecoration(
          border: InputBorder.none,
          hintText: "...",
          hintStyle: TextStyle(color: Colors.grey),
          contentPadding: EdgeInsets.symmetric(vertical: 12)
        ),
      ),
    );
  }
}
