import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart' as ex;
import 'dart:typed_data';

void main() => runApp(const MaterialApp(home: ExcelApp(), debugShowCheckedModeBanner: false));

class ExcelApp extends StatefulWidget {
  const ExcelApp({super.key});
  @override
  State<ExcelApp> createState() => _ExcelAppState();
}

class _ExcelAppState extends State<ExcelApp> {
  List<List<TextEditingController>> _controllers = [];

  @override
  void initState() {
    super.initState();
    _addNewRow();
  }

  void _addNewRow() {
    setState(() {
      _controllers.add(List.generate(4, (_) => TextEditingController()));
    });
  }

  // CHỨC NĂNG LƯU FILE - Đã sửa lỗi CellValue
  Future<void> _exportExcel() async {
    var excel = ex.Excel.createExcel();
    ex.Sheet sheetObject = excel['Sheet1'];

    // Tiêu đề: Phải dùng TextCellValue
    sheetObject.appendRow([
      ex.TextCellValue('Tên Sản Phẩm'),
      ex.TextCellValue('Giá Bán'),
      ex.TextCellValue('Giá Nhập'),
      ex.TextCellValue('Số Lượng'),
    ]);

    // Dữ liệu: Chuyển đổi từ Controller sang TextCellValue
    for (var row in _controllers) {
      sheetObject.appendRow([
        ex.TextCellValue(row[0].text),
        ex.TextCellValue(row[1].text),
        ex.TextCellValue(row[2].text),
        ex.TextCellValue(row[3].text),
      ]);
    }

    excel.save(fileName: "DuLieuBanHang.xlsx");
  }

  // CHỨC NĂNG MỞ FILE
  Future<void> _importExcel() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
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
            // Bỏ qua hàng tiêu đề
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

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Phần mềm Excel 4 Cột - Full Chức Năng'),
        backgroundColor: Colors.blueAccent,
        actions: [
          IconButton(icon: const Icon(Icons.file_open), onPressed: _importExcel),
          IconButton(icon: const Icon(Icons.save), onPressed: _exportExcel),
        ],
      ),
      body: SingleChildScrollView(
        padding: const EdgeInsets.all(15),
        child: Table(
          border: TableBorder.all(color: Colors.grey.shade400),
          columnWidths: const {
            0: FlexColumnWidth(3),
            1: FlexColumnWidth(1.5),
            2: FlexColumnWidth(1.5),
            3: FlexColumnWidth(1.5),
          },
          children: [
            TableRow(
              decoration: const BoxDecoration(color: Colors.blue),
              children: ['Tên SP', 'Giá Bán', 'Giá Nhập', 'Số Lượng'].map((text) => 
                Padding(
                  padding: const EdgeInsets.all(12),
                  child: Text(text, style: const TextStyle(color: Colors.white, fontWeight: FontWeight.bold))
                )
              ).toList(),
            ),
            ..._controllers.map((rowControllers) => TableRow(
              children: rowControllers.map((ctrl) => Padding(
                padding: const EdgeInsets.symmetric(horizontal: 10),
                child: TextField(controller: ctrl, decoration: const InputDecoration(border: InputBorder.none)),
              )).toList(),
            )),
          ],
        ),
      ),
      floatingActionButton: FloatingActionButton(onPressed: _addNewRow, child: const Icon(Icons.add)),
    );
  }
}