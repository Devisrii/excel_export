// ignore_for_file: use_build_context_synchronously

import 'dart:typed_data';
import 'package:flutter/material.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as xls;
import 'package:file_saver/file_saver.dart';
import 'package:file_picker/file_picker.dart';
import 'zip_csv_parser.dart';
// import 'data.dart' as data;

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Excel Export',
      theme: ThemeData(colorScheme: ColorScheme.fromSeed(seedColor: Colors.deepPurple)),
      home: const MyHomePage(title: 'Excel Export Demo'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key, required this.title});
  final String title;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: Text(widget.title)),
      body: Center(
        child: ElevatedButton(
          onPressed: () async {
            final res = await loadReport();
            await exportToExcel(res);
            ScaffoldMessenger.of(
              context,
            ).showSnackBar(const SnackBar(content: Text('Excel exported successfully!. Check your download folder')));
          },
          child: const Text("Excel Export"),
        ),
      ),
    );
  }
}

Future<dynamic> loadReport() async {
  dynamic report;
  // Option A: from a known path
  // report = await ZipCsvParser.parse('GSTR1.zip');
  // return report;

  // Option B: using file_picker
  final result = await FilePicker.platform.pickFiles(type: FileType.custom, allowedExtensions: ['zip']);
  if (result != null) {
    report = await ZipCsvParser.parse(result.files.single.path!);
  } else {
    throw Exception("File not picked");
  }
  return report;
}

Future<void> exportToExcel(Map<String, List<List<String>>> report) async {
  final workbook = xls.Workbook();
  workbook.worksheets.clear();

  xls.HAlignType getAlignType(String text) {
    final t = text.toLowerCase();
    if (t.contains('amount') || t.contains('value') || t.contains('total') || t.contains('rate')) {
      return xls.HAlignType.right;
    }
    return xls.HAlignType.left;
  }

  void addSheet(String sheetName, List<List<String>> data) {
    if (data.isEmpty) return;

    final sheet = workbook.worksheets.addWithName(sheetName);

    // First row = Header
    final headers = data.first;

    // Title
    final titleCell = sheet.getRangeByIndex(1, 1);
    titleCell.setText("Summary for $sheetName");
    sheet.getRangeByIndex(1, 1, 1, headers.length).merge();

    // Headers
    for (int i = 0; i < headers.length; i++) {
      final cell = sheet.getRangeByIndex(4, i + 1);
      cell.setText(headers[i]);
      cell.cellStyle.hAlign = getAlignType(headers[i]);
    }

    // Data rows
    for (int r = 1; r < data.length; r++) {
      final row = data[r];
      for (int c = 0; c < row.length; c++) {
        final cell = sheet.getRangeByIndex(r + 4, c + 1);
        cell.setText(row[c]);
        cell.cellStyle.hAlign = getAlignType(headers[c]);
      }
    }
  }

  // 🔥 Create sheets dynamically
  report.forEach((sheetName, data) {
    addSheet(sheetName, data);
  });

  final bytes = workbook.saveAsStream();
  workbook.dispose();

  await FileSaver.instance.saveFile(
    name: 'GSTR1_${DateTime.now().millisecondsSinceEpoch}',
    bytes: Uint8List.fromList(bytes),
    fileExtension: 'xlsx',
    mimeType: MimeType.microsoftExcel,
  );
}
