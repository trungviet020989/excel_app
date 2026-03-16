import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart' as ex;
import 'dart:typed_data';
import 'dart:io';
import 'package:shared_preferences/shared_preferences.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:share_plus/share_plus.dart';

void main() => runApp(MaterialApp(
      title: 'Quản Lý Kho',
      home: const ExcelApp(),
      debugShowCheckedModeBanner: false,
      theme: ThemeData(primarySwatch: Colors.teal, useMaterial3: true),
    ));

class ExcelApp extends StatefulWidget {
  const ExcelApp({super.key});
  @override
  State<ExcelApp> createState() => _ExcelAppState();
}

class _ExcelAppState extends State<ExcelApp> {
  List<List<TextEditingController>> _controllers = [];
  final List<TextEditingController> _topInputCtrls = List.generate(4, (_) => TextEditingController());
  String? _defaultPath;
  String? _currentFileNameOnly;
  String _searchQuery = "";
  final ScrollController _scrollController = ScrollController();

  @override
  void initState() {
    super.initState();
    _loadDefaultPath();
  }

  @override
  void dispose() {
    _scrollController.dispose();
    for (var ctrl in _topInputCtrls) ctrl.dispose();
    for (var row in _controllers) {
      for (var ctrl in row) ctrl.dispose();
    }
    super.dispose();
  }

  // 1. CÀI ĐẶT THƯ MỤC
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
      _showSnackBar("Đã cài đặt thư mục lưu.");
    }
  }

  // 2. THÊM DÒNG MỚI TỪ THANH CỐ ĐỊNH
  void _addFromTop() {
    if (_topInputCtrls[0].text.isEmpty) {
      _showSnackBar("Vui lòng nhập tên sản phẩm");
      return;
    }

    final newRow = [
      TextEditingController(text: _topInputCtrls[0].text),
      TextEditingController(text: _topInputCtrls[1].text),
      TextEditingController(text: _topInputCtrls[2].text),
      TextEditingController(text: _topInputCtrls[3].text),
    ];

    setState(() {
      _controllers.add(newRow);
      for (var ctrl in _topInputCtrls) ctrl.clear();
    });

    FocusScope.of(context).unfocus();

    Future.microtask(() {
      if (_scrollController.hasClients) {
        _scrollController.animateTo(
          _scrollController.position.maxScrollExtent,
          duration: const Duration(milliseconds: 300),
          curve: Curves.easeOut,
        );
      }
    });
  }

  void _removeRow(int index) {
    setState(() {
      _controllers[index].forEach((c) => c.dispose());
      _controllers.removeAt(index);
    });
  }

  // 3. LƯU FILE EXCEL
  Future<void> _exportExcel() async {
    try {
      if (_defaultPath == null) {
        _showSnackBar("Vui lòng cài đặt thư mục lưu trong Settings.");
        return;
      }
      if (Platform.isAndroid) await [Permission.storage, Permission.manageExternalStorage].request();
      
      var excel = ex.Excel.createExcel();
      ex.Sheet sheetObject = excel['Sheet1'];
      sheetObject.appendRow([ex.TextCellValue('Tên SP'), ex.TextCellValue('Giá Bán'), ex.TextCellValue('Giá Nhập'), ex.TextCellValue('SL')]);
      
      for (var row in _controllers) {
        sheetObject.appendRow([
          ex.TextCellValue(row[0].text),
          ex.TextCellValue(row[1].text),
          ex.TextCellValue(row[2].text),
          ex.TextCellValue(row[3].text)
        ]);
      }

      final List<int>? fileBytes = excel.save();
      if (fileBytes == null) return;
      
      String suggestion = _currentFileNameOnly ?? "San_pham_moi";
      String? customName = await _showFileNameDialog(suggestion);
      if (customName == null || customName.isEmpty) return;
      
      final file = File("$_defaultPath/$customName.xlsx");
      await file.writeAsBytes(Uint8List.fromList(fileBytes), flush: true);
      setState(() => _currentFileNameOnly = customName);
      _showSnackBar("Đã lưu thành công!");
    } catch (e) { _showSnackBar("Lỗi: $e"); }
  }

  // 4. MỞ FILE EXCEL
  Future<void> _importExcel() async {
    try {
      FilePickerResult? result = await FilePicker.platform.pickFiles(
        type: FileType.custom, 
        allowedExtensions: ['xlsx'],
        withData: true
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
            break;
          }
          _showSnackBar("Đã nhập dữ liệu từ $fileName");
        }
      }
    } catch (e) { _showSnackBar("Lỗi khi mở file: $e"); }
  }

  // 5. CHIA SẺ FILE
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

  // 6. GIAO DIỆN
  @override
  Widget build(BuildContext context) {
    List<List<TextEditingController>> filteredRows = _controllers.where((row) {
      return row[0].text.toLowerCase().contains(_searchQuery.toLowerCase());
    }).toList();

    return Scaffold(
      backgroundColor: Colors.blueGrey[50],
      appBar: AppBar(
        title: const Text('QUẢN LÝ KHO', style: TextStyle(color: Colors.white, fontSize: 16, fontWeight: FontWeight.bold)),
        flexibleSpace: Container(decoration: const BoxDecoration(gradient: LinearGradient(colors: [Colors.teal, Colors.green]))),
        actions: [
          IconButton(icon: const Icon(Icons.note_add, color: Colors.white), onPressed: () => setState(() => _controllers.clear())),
          IconButton(icon: const Icon(Icons.share, color: Colors.white), onPressed: _pickAndShareFile),
          IconButton(icon: const Icon(Icons.settings, color: Colors.white), onPressed: _settingsPath),
          IconButton(icon: const Icon(Icons.file_open, color: Colors.white), onPressed: _importExcel),
          IconButton(icon: const Icon(Icons.save, color: Colors.white), onPressed: _exportExcel),
        ],
      ),
      body: Column(
        children: [
          Padding(
            padding: const EdgeInsets.all(8.0),
            child: TextField(
              onChanged: (v) => setState(() => _searchQuery = v),
              decoration: InputDecoration(
                hintText: "Tìm sản phẩm...",
                prefixIcon: const Icon(Icons.search),
                filled: true, fillColor: Colors.white,
                border: OutlineInputBorder(borderRadius: BorderRadius.circular(10), borderSide: BorderSide.none),
                contentPadding: EdgeInsets.zero,
              ),
            ),
          ),
          Container(
            color: Colors.orange[100],
            padding: const EdgeInsets.all(8),
            child: Row(
              children: [
                _buildInput(_topInputCtrls[0], "Tên SP", 3),
                _buildInput(_topInputCtrls[1], "Bán", 2, isNum: true),
                _buildInput(_topInputCtrls[2], "Nhập", 2, isNum: true),
                _buildInput(_topInputCtrls[3], "SL", 1, isNum: true),
                IconButton(icon: const Icon(Icons.add_circle, color: Colors.orange, size: 35), onPressed: _addFromTop),
              ],
            ),
          ),
          Expanded(
            child: Scrollbar(
              controller: _scrollController,
              thumbVisibility: true,
              thickness: 7,
              child: ListView.builder(
                controller: _scrollController,
                itemCount: filteredRows.length,
                itemExtent: 55,
                itemBuilder: (context, index) {
                  return Card(
                    margin: const EdgeInsets.symmetric(horizontal: 8, vertical: 2),
                    child: Row(
                      children: [
                        _buildCell(filteredRows[index][0], 3),
                        _buildCell(filteredRows[index][1], 2, isNum: true),
                        _buildCell(filteredRows[index][2], 2, isNum: true),
                        _buildCell(filteredRows[index][3], 1, isNum: true),
                        IconButton(
                          icon: const Icon(Icons.delete, color: Colors.redAccent, size: 20),
                          onPressed: () => _removeRow(_controllers.indexOf(filteredRows[index])),
                        )
                      ],
                    ),
                  );
                },
              ),
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildInput(TextEditingController c, String h, int f, {bool isNum = false}) => Expanded(flex: f, child: Padding(padding: const EdgeInsets.symmetric(horizontal: 2), child: TextField(controller: c, keyboardType: isNum ? TextInputType.number : TextInputType.text, decoration: InputDecoration(hintText: h, filled: true, fillColor: Colors.white, border: OutlineInputBorder(borderRadius: BorderRadius.circular(5), borderSide: BorderSide.none), contentPadding: const EdgeInsets.symmetric(horizontal: 5)))));
  Widget _buildCell(TextEditingController c, int f, {bool isNum = false}) => Expanded(flex: f, child: TextField(controller: c, keyboardType: isNum ? TextInputType.number : TextInputType.text, textAlign: isNum ? TextAlign.center : TextAlign.left, style: const TextStyle(fontSize: 13), decoration: const InputDecoration(border: InputBorder.none, contentPadding: EdgeInsets.symmetric(horizontal: 5))));
  
  void _showSnackBar(String m) => ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(m)));

  Future<String?> _showFileNameDialog(String initialName) async {
    TextEditingController _nameCtrl = TextEditingController(text: initialName);
    return showDialog<String>(
      context: context,
      builder: (context) => AlertDialog(
        title: const Text("Lưu file Excel"),
        content: TextField(controller: _nameCtrl, decoration: const InputDecoration(suffixText: ".xlsx"), autofocus: true),
        actions: [
          TextButton(onPressed: () => Navigator.pop(context), child: const Text("Hủy")),
          ElevatedButton(onPressed: () => Navigator.pop(context, _nameCtrl.text), child: const Text("Xác nhận")),
        ],
      ),
    );
  }
}
