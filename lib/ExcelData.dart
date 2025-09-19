import 'dart:io';

import 'package:excel/excel.dart';
import 'package:excel_file/const.dart';
import 'package:excel_file/main.dart';

class ExcelData {
  int rowIndex;
  String? name;
  String? branchName;

  ExcelData(this.rowIndex);

  Map<DateTime?, Map<int, int>> quantityByHourAtDateList =
      {}; // date : hour : quantity

  // int get total => (quantity ?? 0) * (parts ?? 1);

  String get category => getCategory(name ?? '') ?? 'غير محدد';

  @override
  String toString() {
    return "$rowIndex\n$branchName\n$name\n\t\t$quantityByHourAtDateList";
  }
}

class SaveExcelData {
  final List<ExcelData> list;

  SaveExcelData(this.list);

  void saveExcel() async {
    final outFile = File('./output.xlsx');
    progress.value = 0;

    Excel excel = Excel.decodeBytes(
      await File('./table_save.xlsx').readAsBytes(),
    );

    progress.value = 0.2;

    final sheet = excel.tables.values.first;

    for (var data in list) {
      for (final dateMap in data.quantityByHourAtDateList.entries) {
        for (var hourQuantityMap in dateMap.value.entries) {
          sheet.appendRow([
            TextCellValue(data.name ?? ''),
            TextCellValue(data.branchName ?? ''),
            DateCellValue.fromDateTime(dateMap.key!),
            IntCellValue(hourQuantityMap.key),
            IntCellValue(hourQuantityMap.value),
            TextCellValue(data.category),
            DoubleCellValue(
              getTotalQuantity(data.name!, hourQuantityMap.value).toDouble(),
            ),
          ]);
        }
      }
    }

    final bytes = excel.encode();
    progress.value = 0.6;
    await outFile.writeAsBytes(bytes!);
    progress.value = 1;
  }
}
