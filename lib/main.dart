import 'dart:typed_data';
import 'package:flutter/material.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as xls;
import 'package:file_saver/file_saver.dart';

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
  final Map<String, List<List<String>>> report = {
    "b2b": [
      [
        "GSTIN",
        "Invoice No",
        "Invoice Type",
        "Date",
        "POS",
        "Rate",
        "Invoice Value",
        "Taxable",
        "CGST",
        "SGST",
        "IGST",
        "Cess",
      ],
      [
        "33MVHPS9823H1ZV",
        "TB242521",
        "Regular B2B",
        "2025-01-25",
        "33-TamilNadu",
        "12.0",
        "525.0",
        "468.74",
        "28.13",
        "28.13",
        "0.0",
        "0.0",
      ],
    ],

    "b2cl": [],

    "b2cs": [
      ["Supply Type", "Type", "POS", "Rate", "Taxable", "CGST", "SGST", "IGST", "Cess", "Total"],
      ["INTRA", "OE", "33-TamilNadu", "18.0", "1210.16", "108.92", "108.92", "0.0", "0.0", "1428.0"],
      ["INTRA", "OE", "33-TamilNadu", "28.0", "1562.5", "218.75", "218.75", "0.0", "0.0", "2000.0"],
    ],

    "nil": [
      ["Supply Type", "Exempt Amount", "Nil Amount", "Non-GST Supply"],
      ["Intra-State supplies to unregistered persons", "0.0", "0", "800.0"],
    ],

    "docs": [
      ["From", "To", "Document Type", "Total Number", "Cancelled"],
      ["VB25261", "VB25263", "Credit Note", "3", "0"],
      ["TB242521", "TB242521", "Invoices for outward supply", "1", "0"],
    ],

    "hsn": [
      ["Description", "UQC", "HSN Code", "Rate", "Taxable", "CGST", "SGST", "IGST", "Cess", "Total", "Qty"],
      ["vicks", "PCS", "123", "12.0", "468.74", "28.13", "28.13", "0.0", "0.0", "525.0", "10.0"],
    ],

    "hsnB2b": [
      ["Description", "UQC", "HSN Code", "Rate", "Taxable", "CGST", "SGST", "IGST", "Cess", "Total", "Qty"],
      ["vicks", "PCS", "123", "12.0", "468.74", "28.13", "28.13", "0.0", "0.0", "525.0", "10.0"],
    ],

    "hsnB2c": [
      ["Description", "UQC", "HSN Code", "Rate", "Taxable", "CGST", "SGST", "IGST", "Cess", "Total", "Qty"],
      ["Demo Pen", "PCS", "03061720", "12.0", "31317.84", "1879.08", "1879.08", "0.0", "0.0", "35076.0", "226.0"],
    ],

    "cdnr": [],
    "cdnur": [],
    "exp": [],
  };

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: Text(widget.title)),
      body: Center(
        child: ElevatedButton(
          onPressed: () async {
            await exportToExcel(report);
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
