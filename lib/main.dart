import 'dart:typed_data';
import 'package:flutter/material.dart';
import 'package:jiffy/jiffy.dart';
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
  final Map<String, dynamic> report = {
    "b2b": [
      {
        "id": 63,
        "gstNo": "33MVHPS9823H1ZV",
        "invNo": "TB242521",
        "invType": "Regular B2B",
        "invDate": "2025-01-25",
        "pos": 33,
        "posName": "TamilNadu",
        "revCharge": "N",
        "taxRatio": 12.0,
        "invAmt": 525.0,
        "taxable": 468.74,
        "cgst": 28.13,
        "sgst": 28.13,
        "igst": 0.0,
        "cess": 0.0,
        "total": 525.0,
      },
    ],

    "b2cl": [],

    "b2cs": [
      {
        "supplyType": "INTRA",
        "type": "OE",
        "pos": 33,
        "posName": "TamilNadu",
        "taxRatio": 18.0,
        "taxable": 1210.16,
        "cgst": 108.92,
        "sgst": 108.92,
        "igst": 0.0,
        "cess": 0.0,
        "total": 1428.0,
      },
      {
        "supplyType": "INTRA",
        "type": "OE",
        "pos": 33,
        "posName": "TamilNadu",
        "taxRatio": 28.0,
        "taxable": 1562.5,
        "cgst": 218.75,
        "sgst": 218.75,
        "igst": 0.0,
        "cess": 0.0,
        "total": 2000.0,
      },
    ],

    "nil": [
      {"supplyType": "Intra-State supplies to unregistered persons", "exptAmt": 0.0, "nilAmt": 0, "ngsupAmt": 800.0},
    ],

    "docs": [
      {"from": "VB25261", "to": "VB25263", "docTyp": "Credit Note", "totnum": 3, "cancel": 0},
      {"from": "TB242521", "to": "TB242521", "docTyp": "Invoices for outward supply", "totnum": 1, "cancel": 0},
    ],

    "hsn": [
      {
        "description": "vicks",
        "uqc": {"displayText": "PCS"},
        "hsnSacCode": "123",
        "taxRatio": 12.0,
        "taxable": 468.74,
        "cgst": 28.13,
        "sgst": 28.13,
        "igst": 0.0,
        "cess": 0.0,
        "total": 525.0,
        "qty": 10.0,
      },
    ],

    "hsnB2b": [
      {
        "description": "vicks",
        "uqc": {"displayText": "PCS"},
        "hsnSacCode": "123",
        "taxRatio": 12.0,
        "taxable": 468.74,
        "cgst": 28.13,
        "sgst": 28.13,
        "igst": 0.0,
        "cess": 0.0,
        "total": 525.0,
        "qty": 10.0,
      },
    ],

    "hsnB2c": [
      {
        "description": "Demo Pen",
        "uqc": {"displayText": "PCS"},
        "hsnSacCode": "03061720",
        "taxRatio": 12.0,
        "taxable": 31317.84,
        "cgst": 1879.08,
        "sgst": 1879.08,
        "igst": 0.0,
        "cess": 0.0,
        "total": 35076.0,
        "qty": 226.0,
      },
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

Future<void> exportToExcel(Map<String, dynamic> report) async {
  final workbook = xls.Workbook();
  workbook.worksheets.clear();

  List getList(String key) => (report[key] as List?) ?? [];

  String formatDate(String? date) {
    if (date == null || date.isEmpty) return '';
    return Jiffy.parse(date, pattern: 'yyyy-MM-dd').format(pattern: 'dd-MMM-yyyy');
  }

  xls.HAlignType getAlignType(String text) {
    final t = text.toLowerCase();
    if (t.contains('amount') || t.contains('value') || t.contains('total') || t.contains('rate')) {
      return xls.HAlignType.right;
    }
    return xls.HAlignType.left;
  }

  xls.Worksheet addSheet(String name, String title, List<String> headers, List<List<dynamic>> rows) {
    final sheet = workbook.worksheets.addWithName(name);

    final titleCell = sheet.getRangeByIndex(1, 1);
    titleCell.setText(title);
    sheet.getRangeByIndex(1, 1, 1, headers.length).merge();

    for (var i = 0; i < headers.length; i++) {
      final cell = sheet.getRangeByIndex(4, i + 1);
      cell.setText(headers[i]);
      cell.cellStyle.hAlign = getAlignType(headers[i]);
    }

    for (var r = 0; r < rows.length; r++) {
      for (var c = 0; c < rows[r].length; c++) {
        final cell = sheet.getRangeByIndex(r + 5, c + 1);
        cell.setText(rows[r][c]?.toString() ?? '');
        cell.cellStyle.hAlign = getAlignType(headers[c]);
      }
    }

    return sheet;
  }

  // ---------------------- B2B ----------------------
  addSheet(
    'b2b,sez,de',
    'Summary for b2b',
    [
      'GSTIN/UIN of Recipient',
      'Receiver Name',
      'Invoice Number',
      'Invoice date',
      'Invoice Value',
      'Place Of Supply',
      'Reverse Charge',
      'Invoice Type',
      'Rate',
      'Taxable Value',
      'Cess Amount',
    ],
    getList('b2b').map<List<dynamic>>((el) {
      return [
        el['gstNo'],
        el['receiverName'] ?? '',
        el['invNo'],
        formatDate(el['invDate']),
        el['invAmt'],
        '${el['pos']}-${el['posName']}',
        el['revCharge'],
        el['invType'],
        el['taxRatio'],
        el['taxable'],
        el['cess'],
      ];
    }).toList(),
  );

  // ---------------------- B2CS ----------------------
  addSheet(
    'b2cs',
    'Summary for b2cs',
    ['Type', 'Place Of Supply', 'Rate', 'Taxable Value', 'Cess Amount'],
    getList('b2cs').map<List<dynamic>>((el) {
      return [el['type'], '${el['pos']}-${el['posName']}', el['taxRatio'], el['taxable'], el['cess']];
    }).toList(),
  );

  // ---------------------- B2CL ----------------------
  addSheet(
    'b2cl',
    'Summary for b2cl',
    ['Invoice Number', 'Invoice date', 'Invoice Value', 'Place Of Supply', 'Rate', 'Taxable Value', 'Cess Amount'],
    getList('b2cl').map<List<dynamic>>((el) {
      return [
        el['invNo'],
        formatDate(el['invDate']),
        el['invAmt'],
        '${el['pos']}-${el['posName']}',
        el['taxRatio'],
        el['taxable'],
        el['cess'],
      ];
    }).toList(),
  );

  // ---------------------- EXPORT ----------------------
  addSheet(
    'exp',
    'Summary for Export',
    ['Export Type', 'Invoice Number', 'Invoice date', 'Invoice Value', 'Rate', 'Taxable Value', 'Cess Amount'],
    getList('exp').map<List<dynamic>>((el) {
      return [
        el['invType'],
        el['invNo'],
        formatDate(el['invDate']),
        el['invAmt'],
        el['taxRatio'],
        el['taxable'],
        el['cess'],
      ];
    }).toList(),
  );

  // ---------------------- HSN B2B ----------------------
  addSheet(
    'hsn(b2b)',
    'Summary for HSN(b2b)',
    [
      'HSN',
      'Description',
      'UQC',
      'Total Quantity',
      'Total Value',
      'Rate',
      'Taxable Value',
      'IGST',
      'CGST',
      'SGST',
      'Cess',
    ],
    getList('hsnB2b').map<List<dynamic>>((el) {
      return [
        el['hsnSacCode'],
        el['description'],
        el['uqc']?['displayText'],
        el['qty'],
        el['total'],
        el['taxRatio'],
        el['taxable'],
        el['igst'],
        el['cgst'],
        el['sgst'],
        el['cess'],
      ];
    }).toList(),
  );

  // ---------------------- HSN B2C ----------------------
  addSheet(
    'hsn(b2c)',
    'Summary for HSN(b2c)',
    [
      'HSN',
      'Description',
      'UQC',
      'Total Quantity',
      'Total Value',
      'Rate',
      'Taxable Value',
      'IGST',
      'CGST',
      'SGST',
      'Cess',
    ],
    getList('hsnB2c').map<List<dynamic>>((el) {
      return [
        el['hsnSacCode'],
        el['description'],
        el['uqc']?['displayText'],
        el['qty'],
        el['total'],
        el['taxRatio'],
        el['taxable'],
        el['igst'],
        el['cgst'],
        el['sgst'],
        el['cess'],
      ];
    }).toList(),
  );

  // ---------------------- DOCS ----------------------
  addSheet(
    'docs',
    'Summary for document issued',
    ['Nature of Document', 'From', 'To', 'Total', 'Cancelled'],
    getList('docs').map<List<dynamic>>((el) {
      return [el['docTyp'], el['from'], el['to'], el['totnum'], el['cancel']];
    }).toList(),
  );

  // ---------------------- CDNR ----------------------
  addSheet(
    'cdnr',
    'Summary for cdnr',
    ['GSTIN', 'Receiver', 'Note No', 'Date', 'Type', 'POS', 'Value', 'Rate', 'Taxable', 'Cess'],
    getList('cdnr').map<List<dynamic>>((el) {
      return [
        el['gstNo'],
        el['partyName'],
        el['noteNo'],
        formatDate(el['noteDate']),
        el['noteType'],
        '${el['pos']}-${el['posName']}',
        el['noteAmt'],
        el['taxRatio'],
        el['taxable'],
        el['cess'],
      ];
    }).toList(),
  );

  // ---------------------- CDNUR ----------------------
  addSheet(
    'cdnur',
    'Summary for cdnur',
    ['Type', 'Note No', 'Date', 'Note Type', 'POS', 'Value', 'Rate', 'Taxable', 'Cess'],
    getList('cdnur').map<List<dynamic>>((el) {
      return [
        el['supplyType'],
        el['noteNo'],
        formatDate(el['noteDate']),
        el['noteType'],
        '${el['pos']}-${el['posName']}',
        el['noteAmt'],
        el['taxRatio'],
        el['taxable'],
        el['cess'],
      ];
    }).toList(),
  );

  final bytes = workbook.saveAsStream();
  workbook.dispose();

  await FileSaver.instance.saveFile(
    name: 'GSTR1_${DateTime.now().millisecondsSinceEpoch}',
    bytes: Uint8List.fromList(bytes),
    fileExtension: 'xlsx',
    mimeType: MimeType.microsoftExcel,
  );
}
