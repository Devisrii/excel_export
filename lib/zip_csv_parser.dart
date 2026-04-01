import 'dart:io';
import 'package:archive/archive_io.dart';
import 'package:csv/csv.dart';

class ZipCsvParser {
  /// Reads a ZIP file and converts all CSVs inside it into:
  /// Map<fileName (without .csv), List<List<String>>>
  static Future<Map<String, List<List<String>>>> parse(String zipFilePath) async {
    final Map<String, List<List<String>>> report = {};

    final bytes = await File(zipFilePath).readAsBytes();
    final archive = ZipDecoder().decodeBytes(bytes);

    for (final file in archive) {
      if (!file.isFile) continue;

      // Get the file name without extension
      final fullName = file.name.split('/').last; // handle subdirectory paths
      if (!fullName.toLowerCase().endsWith('.csv')) continue;

      final key = fullName.replaceAll(RegExp(r'\.csv$', caseSensitive: false), '');
      final content = String.fromCharCodes(file.content as List<int>);

      if (content.trim().isEmpty) {
        report[key] = [];
        continue;
      }

      final rows = const CsvToListConverter(
        eol: '\n',
      ).convert(content).map((row) => row.map((cell) => cell.toString()).toList()).toList();

      report[key] = rows;
    }

    return report;
  }
}
