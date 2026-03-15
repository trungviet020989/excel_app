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
        primarySwatch: Colors.teal,
        useMaterial3: true,
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
  String _searchQuery = ""; // Biến lưu từ khóa tìm kiếm

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
      _showSnackBar("Vui lòng cài đặt thư mục lưu.");
      return;
    }
    final directory = Directory(_defaultPath!);
    if (!await directory.exists()) return;
    
    List<FileSystemEntity> files = directory.listSync()
        .where((file) => file.path.endsWith('.xlsx'))
        .toList();

    if (files.isEmpty) {
      _showSnackBar("Không tìm thấy file nào.");
      return;
    }

    showDialog(
      context: context,
      builder: (context) => AlertDialog(
        title: const Text("Chọn file gửi đi", style: TextStyle(color: Colors.teal, fontWeight: FontWeight.bold)),
        content: SizedBox(
          width: double.maxFinite,
          child: ListView.builder(
            shrinkWrap: true,
            itemCount: files.length,
            itemBuilder: (context, index) {
              String fileName = files[index].path.split('/').last;
              return Card(
                child: ListTile(
                  leading: const Icon(Icons.file_present, color: Colors.teal),
                  title: Text(fileName),
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

  // --- LOGIC LƯU VÀ MỞ FILE ---
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
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(message), backgroundColor: Colors.teal));
  }

  @override
  Widget build(BuildContext context) {
    bool isKeyboardVisible = MediaQuery.of(context).viewInsets.bottom != 0;

    // Logic lọc danh sách theo tên SP
    List<List<TextEditingController>> filteredRows = _controllers.where((row) {
      return row[0].text.toLowerCase().contains(_searchQuery.toLowerCase());
    }).toList();

    return Scaffold(
      backgroundColor: Colors.blueGrey[50],
      appBar: AppBar(
        title: const Text('QUẢN LÝ KHO', style: TextStyle(color: Colors.white, fontSize: 18, fontWeight: FontWeight.bold)),
        centerTitle: true,
        flexibleSpace: Container(decoration: const BoxDecoration(gradient: LinearGradient(colors: [Colors.teal, Colors.green]))),
        actions: [
          IconButton(
            icon: const Icon(Icons.note_add, color: Colors.white),
            onPressed: () => setState(() {
              _controllers = [List.generate(4, (_) => TextEditingController())];
              _currentFileNameOnly = null;
              _showSnackBar("Đã tạo trang mới");
            }),
          ),
          IconButton(icon: const Icon(Icons.share, color: Colors.white), onPressed: _pickAndShareFile),
          IconButton(icon: const Icon(Icons.settings, color: Colors.white), onPressed: _settingsPath),
          IconButton(icon: const Icon(Icons.file_open, color: Colors.white), onPressed: _importExcel),
          IconButton(icon: const Icon(Icons.save, color: Colors.white), onPressed: _exportExcel),
        ],
      ),
      body: Column(
        children: [
          // 1. Ô TÌM KIẾM SẢN PHẨM
          Container(
            padding: const EdgeInsets.all(10),
            color: Colors.white,
            child: TextField(
              onChanged: (value) => setState(() => _searchQuery = value),
              decoration: InputDecoration(
                hintText: "Tìm tên sản phẩm...",
                prefixIcon: const Icon(Icons.search, color: Colors.teal),
                contentPadding: const EdgeInsets.symmetric(vertical: 0),
                border: OutlineInputBorder(borderRadius: BorderRadius.circular(25), borderSide: BorderSide.none),
                filled: true,
                fillColor: Colors.blueGrey[50],
              ),
            ),
          ),
          // 2. HIỂN THỊ TÊN FILE
          Container(
            width: double.infinity,
            padding: const EdgeInsets.symmetric(vertical: 5, horizontal: 15),
            child: Text(
              _currentFileNameOnly == null ? "🆕 Tệp mới" : "📂 File: $_currentFileNameOnly.xlsx",
              style: const TextStyle(color: Colors.blueGrey, fontSize: 11, fontWeight: FontWeight.bold),
            ),
          ),
          // 3. BẢNG DỮ LIỆU
          Expanded(
            child: SingleChildScrollView(
              padding: const EdgeInsets.symmetric(horizontal: 10),
              child: Card(
                elevation: 2,
                child: Table(
                  columnWidths: const {0: FlexColumnWidth(2), 1: FlexColumnWidth(1.2), 2: FlexColumnWidth(1.2), 3: FlexColumnWidth(0.8)},
                  border: TableBorder.all(color: Colors.teal.shade50),
                  children: [
                    TableRow(
                      decoration: const BoxDecoration(color: Colors.teal),
                      children: ['Tên SP', 'Giá Bán', 'Giá Nhập', 'SL'].map((t) => Padding(
                        padding: const EdgeInsets.all(10), 
                        child: Text(t, style: const TextStyle(color: Colors.white, fontWeight: FontWeight.bold, fontSize: 12))
                      )).toList(),
                    ),
                    ...filteredRows.map((row) => TableRow(
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
        ],
      ),
      floatingActionButton: isKeyboardVisible ? null : FloatingActionButton(
        onPressed: _addNewRow,
        backgroundColor: Colors.teal,
        child: const Icon(Icons.add, color: Colors.white),
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
        decoration: const InputDecoration(border: InputBorder.none, hintText: "...", contentPadding: EdgeInsets.symmetric(vertical: 10)),
      ),
    );
  }
}
