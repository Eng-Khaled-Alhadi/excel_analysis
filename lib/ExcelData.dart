import 'dart:io';

import 'package:excel/excel.dart';
import 'package:excel_file/Firebase_dataBase.dart';
import 'package:excel_file/main.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/services.dart';
import 'package:flutter/foundation.dart';
import 'package:intl/intl.dart';
import 'package:open_file/open_file.dart';
import 'package:collection/collection.dart';

class ExcelData {
  int rowIndex;
  String? name;
  String? branchName;

  ExcelData(this.rowIndex);

  Map<DateTime?, Map<int, int>> quantityByHourAtDateList =
      {}; // date : hour : quantity

  // تحسين: استخدام الكاش المحلي بدلاً من البحث في القائمة في كل مرة
  static final Map<String, String?> _categoryCache = {};
  static final Map<String, num?> _totalQuantityCache = {};

  String? get category => getCategory(name ?? '');

  @override
  String toString() {
    return "$rowIndex\n$branchName\n$name\n\t\t$quantityByHourAtDateList";
  }

  static String? getCategory(String productName) {
    try {
      // التحقق من الكاش أولاً
      if (_categoryCache.containsKey(productName)) {
        return _categoryCache[productName];
      }

      final product = FB.instance.products.firstWhereOrNull(
        (p) => p.name.trim() == productName.trim(),
      );

      // حفظ في الكاش
      _categoryCache[productName] = product?.category;

      return product?.category;
    } catch (e) {
      print('خطأ في getCategory: $e');
      return null;
    }
  }

  static num? getTotalQuantity(String productName, int quantity) {
    try {
      // إنشاء مفتاح كاش فريد
      final cacheKey = '$productName-$quantity';

      // التحقق من الكاش أولاً
      if (_totalQuantityCache.containsKey(cacheKey)) {
        return _totalQuantityCache[cacheKey];
      }

      final product = FB.instance.products.firstWhereOrNull(
        (p) => p.name.trim() == productName.trim(),
      );

      if (product == null) {
        print('المنتج غير موجود: $productName');
        _totalQuantityCache[cacheKey] = null;
        return null;
      }

      final result = product.parts * quantity;

      // حفظ في الكاش
      _totalQuantityCache[cacheKey] = result;

      return result;
    } catch (e) {
      print('خطأ في getTotalQuantity: $e');
      return null;
    }
  }

  // دالة لمسح الكاش
  static void clearCache() {
    _categoryCache.clear();
    _totalQuantityCache.clear();
  }
}

class SaveExcelData {
  final List<ExcelData> list;

  SaveExcelData(this.list);

  void saveExcel() async {
    try {
      progress.value = null;

      final outputFilePath = prefs.getString(outFilePath);

      Excel excel;
      Sheet sheet;

      if (outputFilePath != null && await File(outputFilePath).exists()) {
        try {
          // استخدام الملف الموجود
          final bytes = await File(outputFilePath).readAsBytes();
          excel = Excel.decodeBytes(bytes);

          // البحث عن ورقة "Data" أو أول ورقة متاحة
          if (excel.tables.containsKey("Data")) {
            sheet = excel.tables["Data"]!;
            print('تم العثور على ورقة "Data"');
          } else if (excel.tables.isNotEmpty) {
            final firstSheetName = excel.tables.keys.first;
            sheet = excel.tables[firstSheetName]!;
            print('استخدام الورقة: $firstSheetName');
          } else {
            // إنشاء ورقة جديدة
            sheet = excel['Data'];
            _addHeader(sheet);
          }
        } catch (e) {
          print('خطأ في قراءة الملف الموجود، سيتم إنشاء ملف جديد: $e');
          excel = Excel.createExcel();
          sheet = excel['Data'];
          _addHeader(sheet);
        }
      } else {
        // إنشاء ملف جديد
        excel = Excel.createExcel();
        sheet = excel['Data'];
        _addHeader(sheet);
      }

      progress.value = 0.2;

      final totalProgress = list.fold(
        0,
        (previousValue, element) =>
            previousValue +
            element.quantityByHourAtDateList.values.fold(
              0,
              (previousValue, element) => previousValue + element.length,
            ),
      );

      if (totalProgress == 0) {
        throw Exception('لا توجد بيانات للحفظ');
      }

      int progressLocal = 0;
      int skippedRows = 0;
      int addedRows = 0;

      // تحسين: إعداد قائمة الصفوف مسبقاً ثم إضافتها دفعة واحدة
      List<List<CellValue?>> rowsToAdd = [];

      for (var data in list) {
        // التحقق من البيانات الأساسية
        if (data.name == null || data.name!.isEmpty) {
          continue;
        }

        for (final dateMap in data.quantityByHourAtDateList.entries) {
          if (dateMap.key == null) {
            continue;
          }

          for (var hourQuantityMap in dateMap.value.entries) {
            try {
              final totalQuantity = ExcelData.getTotalQuantity(
                data.name!,
                hourQuantityMap.value,
              );
              final category = data.category;

              if (totalQuantity == null || category == null) {
                skippedRows++;
                print(
                  'تم تخطي الصف: ${data.name} - السبب: ${totalQuantity == null ? "الكمية الإجمالية null" : "التصنيف null"}',
                );
                progressLocal++;
                continue;
              }

              // تحضير الصف للإضافة
              rowsToAdd.add([
                TextCellValue(category),
                TextCellValue(data.name ?? ''),
                TextCellValue(data.branchName ?? ''),
                TextCellValue(DateFormat('EEEE', 'ar').format(dateMap.key!)),
                DateCellValue.fromDateTime(dateMap.key!),
                IntCellValue(hourQuantityMap.key),
                IntCellValue(hourQuantityMap.value),
                DoubleCellValue(totalQuantity.toDouble()),
              ]);

              addedRows++;
            } catch (e) {
              print('خطأ في تحضير الصف: $e');
              skippedRows++;
            }

            progressLocal++;

            // تحديث التقدم كل 50 صف بدلاً من 10 لتحسين الأداء
            if (progressLocal % 50 == 0) {
              progress.value = (progressLocal / totalProgress);
              await Future.delayed(const Duration(microseconds: 1));
            }
          }
        }
      }

      // إضافة جميع الصفوف دفعة واحدة
      for (var row in rowsToAdd) {
        sheet.appendRow(row);
      }

      print('تمت إضافة $addedRows صف');
      if (skippedRows > 0) {
        print('تم تخطي $skippedRows صف');
      }

      if (addedRows == 0) {
        throw Exception(
          'لم يتم إضافة أي بيانات. تأكد من أن المنتجات موجودة في قاعدة البيانات.',
        );
      }

      progress.value = 0.95;

      // حفظ الملف
      final bytes = excel.save();

      if (bytes == null || bytes.isEmpty) {
        throw Exception('فشل في حفظ الملف - البيانات فارغة');
      }

      String? savedPath;

      if (outputFilePath != null) {
        try {
          await File(outputFilePath).writeAsBytes(bytes);
          savedPath = outputFilePath;
          print('تم الحفظ في: $savedPath');
        } catch (e) {
          print('خطأ في الحفظ في المسار المحدد: $e');
          // المحاولة بطريقة بديلة
          savedPath = await FilePicker.platform.saveFile(
            fileName:
                'quantity_report_${DateFormat('yyyyMMdd_HHmmss').format(DateTime.now())}.xlsx',
            bytes: Uint8List.fromList(bytes),
          );
        }
      } else {
        savedPath = await FilePicker.platform.saveFile(
          fileName:
              'quantity_report_${DateFormat('yyyyMMdd_HHmmss').format(DateTime.now())}.xlsx',
          bytes: Uint8List.fromList(bytes),
        );
      }

      if (savedPath == null) {
        throw Exception('لم يتم تحديد مسار الحفظ');
      }

      await Future.delayed(const Duration(milliseconds: 300));
      progress.value = null;

      print('تم حفظ الملف بنجاح: $savedPath');

      await openFile(savedPath);

      // مسح الكاش بعد الانتهاء
      ExcelData.clearCache();
    } catch (e) {
      progress.value = null;
      print('خطأ في saveExcel: $e');
      rethrow;
    }
  }

  void _addHeader(Sheet sheet) {
    try {
      sheet.appendRow([
        TextCellValue('التصنيف'),
        TextCellValue('اسم المنتج'),
        TextCellValue('الفرع'),
        TextCellValue('اليوم'),
        TextCellValue('التاريخ'),
        TextCellValue('الساعة'),
        TextCellValue('الكمية'),
        TextCellValue('الإجمالي'),
      ]);
      print('تم إضافة الهيدر');
    } catch (e) {
      print('خطأ في إضافة الهيدر: $e');
    }
  }

  Future openFile(String file) async {
    try {
      final result = await OpenFile.open(file);
      if (result.type != ResultType.done) {
        print('خطأ في فتح الملف: ${result.message}');
      }
    } catch (e) {
      print('خطأ في openFile: $e');
    }
  }
}
