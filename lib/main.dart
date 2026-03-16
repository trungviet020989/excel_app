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
      theme: ThemeData(primarySwatch: Colors.teal, useMaterial3: true),
    ));

class ExcelApp extends StatefulWidget {
  const ExcelApp({super.key});
  @override
  State<ExcelApp> createState() => _ExcelAppState();
}

class _ExcelAppState extends State<ExcelApp> {
  List<List<TextEditingController>> _controllers = [];
  // Controller cho dòng nhập liệu cố định ở trên
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
    for (var ctrl in _topInputCtrls) {
      ctrl.dispose();
    }
    for (var row in _controllers) {
      for (var ctrl in row) {
        ctrl.dispose();
      }
    }
    super.dispose();
  }

  // Lấy đường dẫn lưu file
  Future<void> _loadDefaultPath() async {
    final prefs = await SharedPreferences.getInstance();
    setState(() => _defaultPath = prefs.getString('default_path'));
  }

  // Thêm dòng mới từ thanh nhập liệu cố định
  void _addFromTop() {
    // Nếu cả 4 ô đều trống thì không thêm
    if (_topInputCtrls.every((c) => c.text.isEmpty)) return;

    final newRow = [
      TextEditingController(text: _topInputCtrls[0].text),
      TextEditingController(text: _topInputCtrls[1].text),
      TextEditingController(text: _topInputCtrls[2].text),
      TextEditingController(text: _topInputCtrls[3].text),
    ];

    setState(() {
      _controllers.add(newRow);
      // Xóa sạch dòng nhập phía trên để chuẩn bị cho sản phẩm tiếp theo
      for (var ctrl in _topInputCtrls) {
        ctrl.clear();
      }
    });

    // Cuộn xuống dòng cuối cùng vừa thêm
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

  // Xóa một dòng bất kỳ
  void _removeRow(int index) {
    setState(() {
      _controllers[index].forEach((c) => c.dispose());
      _controllers.removeAt(index);
    });
  }

  // --- CÁC HÀM XỬ LÝ FILE (GIỮ NGUYÊN) ---
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
      
      final file = File("$_defaultPath/${customFileName.endsWith('.xlsx') ? customFileName : '$customFileName.xlsx'}");
      await file.writeAsBytes(bytes, flush: true);
      setState(() => _currentFileNameOnly = customFileName.replaceAll('.xlsx', ''));
      _showSnackBar("Đã lưu thành công!");
    } catch (e) { _showSnackBar("Lỗi lưu file: $e"); }
  }

  // --- GIAO DIỆN ---
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
          // 1. THANH TÌM KIẾM
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 5),
            color: Colors.white,
            child: TextField(
              onChanged: (v) => setState(() => _searchQuery = v),
              decoration: InputDecoration(
                hintText: "Tìm tên sản phẩm...",
                prefixIcon: const Icon(Icons.search, size: 20),
                filled: true,
                fillColor: Colors.grey[100],
                border: OutlineInputBorder(borderRadius: BorderRadius.circular(10), borderSide: BorderSide.none),
                contentPadding: EdgeInsets.zero,
              ),
            ),
          ),

          // 2. DÒNG NHẬP LIỆU CỐ ĐỊNH (STICKY INPUT)
          Container(
            color: Colors.orange[50],
            padding: const EdgeInsets.symmetric(vertical: 8, horizontal: 5),
            child: Row(
              children: [
                _buildInputCell(_topInputCtrls[0], "Tên SP", flex: 20),
                _buildInputCell(_topInputCtrls[1], "Bán", flex: 12, isNum: true),
                _buildInputCell(_topInputCtrls[2], "Nhập", flex: 12, isNum: true),
                _buildInputCell(_topInputCtrls[3], "SL", flex: 8, isNum: true),
                IconButton(
                  icon: const Icon(Icons.add_circle, color: Colors.orange, size: 30),
                  onPressed: _addFromTop,
                )
              ],
            ),
          ),

          // 3. DANH SÁCH DỮ LIỆU
          Expanded(
            child: Scrollbar(
              controller: _scrollController,
              thumbVisibility: true, // Luôn hiện thanh cuộn bên phải
              thickness: 8,
              radius: const Radius.circular(10),
              child: ListView.builder(
                controller: _scrollController,
                padding: const EdgeInsets.all(10),
                itemCount: filteredRows.length,
                itemExtent: 50, // Cố định chiều cao dòng để mượt tuyệt đối
                itemBuilder: (context, index) {
                  return Card(
                    margin: const EdgeInsets.only(bottom: 2),
                    elevation: 0,
                    shape: RoundedRectangleBorder(side: BorderSide(color: Colors.grey[200]!)),
                    child: Row(
                      children: [
                        _buildDataCell(filteredRows[index][0], flex: 20),
                        _buildDataCell(filteredRows[index][1], flex: 12, isNum: true),
                        _buildDataCell(filteredRows[index][2], flex: 12, isNum: true),
                        _buildDataCell(filteredRows[index][3], flex: 8, isNum: true),
                        IconButton(
                          icon: const Icon(Icons.delete_outline, color: Colors.redAccent, size: 20),
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

  // Ô nhập cho dòng cố định
  Widget _buildInputCell(TextEditingController ctrl, String hint, {required int flex, bool isNum = false}) {
    return Expanded(
      flex: flex,
      child: Container(
        margin: const EdgeInsets.symmetric(horizontal: 2),
        padding: const EdgeInsets.symmetric(horizontal: 4),
        decoration: BoxDecoration(color: Colors.white, borderRadius: BorderRadius.circular(5)),
        child: TextField(
          controller: ctrl,
          keyboardType: isNum ? TextInputType.number : TextInputType.text,
          style: const TextStyle(fontSize: 13),
          decoration: InputDecoration(hintText: hint, border: InputBorder.none, hintStyle: const TextStyle(fontSize: 11)),
        ),
      ),
    );
  }

  // Ô hiển thị dữ liệu trong danh sách
  Widget _buildDataCell(TextEditingController ctrl, {required int flex, bool isNum = false}) {
    return Expanded(
      flex: flex,
      child: TextField(
        controller: ctrl,
        keyboardType: isNum ? TextInputType.number : TextInputType.text,
        textAlign: isNum ? TextAlign.center : TextAlign.left,
        style: const TextStyle(fontSize: 13),
        decoration: const InputDecoration(border: InputBorder.none, contentPadding: EdgeInsets.symmetric(horizontal: 5)),
      ),
    );
  }

  // --- CÁC HÀM TIỆN ÍCH KHÁC ---
  String _suggestNextFileName() { /* ... như cũ ... */ return "Sản phẩm"; }
  Future<void> _pickAndShareFile() async { /* ... như cũ ... */ }
  Future<void> _settingsPath() async { /* ... như cũ ... */ }
  Future<void> _importExcel() async { /* ... như cũ ... */ }
  void _showSnackBar(String msg) { ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(msg))); }
  Future<String?> _showFileNameDialog(String initialName) async { 
    TextEditingController c = TextEditingController(text: initialName);
    return showDialog<String>(context: context, builder: (ctx) => AlertDialog(
      title: const Text("Lưu file"), content: TextField(controller: c),
      actions: [TextButton(onPressed: () => Navigator.pop(ctx, c.text), child: const Text("Lưu"))],
    ));
  }
}
