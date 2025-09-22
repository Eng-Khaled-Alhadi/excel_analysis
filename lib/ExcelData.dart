import 'dart:io';

import 'package:excel/excel.dart';
import 'package:excel_file/const.dart';
import 'package:excel_file/main.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/services.dart';
import 'package:intl/intl.dart';
import 'package:open_file/open_file.dart';

class ExcelData {
  int rowIndex;
  String? name;
  String? branchName;

  ExcelData(this.rowIndex);

  Map<DateTime?, Map<int, int>> quantityByHourAtDateList =
      {}; // date : hour : quantity

  // int get total => (quantity ?? 0) * (parts ?? 1);

  String? get category => getCategory(name ?? '');

  @override
  String toString() {
    return "$rowIndex\n$branchName\n$name\n\t\t$quantityByHourAtDateList";
  }
}

class SaveExcelData {
  final List<ExcelData> list;

  SaveExcelData(this.list);

  void saveExcel() async {
    // final outFile = File('./output.xlsx');
    progress.value = null;

    final outputFilePath = prefs.getString(outFilePath);

    Excel excel = Excel.decodeBytes(
      outputFilePath != null ? await File(outputFilePath).readAsBytes() : (await rootBundle.load('assets/table_save.xlsx')).buffer.asUint8List(),
    );

    progress.value = 0.2;

    final sheet = excel.tables.values.first;

    final totalProgress = list.fold(
      0,
      (previousValue, element) =>
          previousValue +
          element.quantityByHourAtDateList.values.fold(
            0,
            (previousValue, element) => previousValue + element.length,
          ),
    );
    int progressLocal = 0;


    for (var data in list) {
      for (final dateMap in data.quantityByHourAtDateList.entries) {
        for (var hourQuantityMap in dateMap.value.entries) {
          final totalQuantity = getTotalQuantity(
            data.name!,
            hourQuantityMap.value,
          );
          final category = data.category;
          if (totalQuantity == null || category == null) {
            continue;
          }
          sheet.appendRow([
            TextCellValue(data.category!),
            TextCellValue(data.name ?? ''),
            TextCellValue(data.branchName ?? ''),
            TextCellValue(DateFormat('EEEE','ar').format(dateMap.key!)),
            DateCellValue.fromDateTime(dateMap.key!),
            IntCellValue(hourQuantityMap.key),
            IntCellValue(hourQuantityMap.value),
            DoubleCellValue(
              getTotalQuantity(data.name!, hourQuantityMap.value)!.toDouble(),
            ),
          ]);
          progress.value = (progressLocal++ / totalProgress);
        }
        await Future.delayed(const Duration(microseconds: 1));
      }
    }

    progress.value = 1;
    final bytes = excel.encode();
    print("$outputFilePath ----> 0");
    final path = outputFilePath != null ? (await File(outputFilePath).writeAsBytes(bytes!).then((value) => value.path)) : await FilePicker.platform.saveFile(fileName: 'quantity_report.xlsx', bytes: Uint8List.fromList(bytes!));
    // ;
    await Future.delayed(const Duration(milliseconds: 300));
    progress.value = null;
    await openFile(path!);
  }

  Future openFile(String file) async {
    await OpenFile.open(file);
  }
}
