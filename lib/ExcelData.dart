import 'dart:io';

import 'package:excel/excel.dart';
import 'package:excel_file/const.dart';
import 'package:excel_file/main.dart';
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
    final outFile = File('./output.xlsx');
    progress.value = null;

    Excel excel = Excel.decodeBytes(
      await File('./table_save.xlsx').readAsBytes(),
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
            TextCellValue(data.name ?? ''),
            TextCellValue(data.branchName ?? ''),
            DateCellValue.fromDateTime(dateMap.key!),
            IntCellValue(hourQuantityMap.key),
            IntCellValue(hourQuantityMap.value),
            TextCellValue(data.category!),
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
    final bytes = excel.save();
    await outFile.writeAsBytes(bytes!);
    await Future.delayed(const Duration(milliseconds: 300));
    progress.value = null;
    await openFile(outFile);
  }

  Future openFile(File file) async {
    await OpenFile.open(file.absolute.path);
  }
}
