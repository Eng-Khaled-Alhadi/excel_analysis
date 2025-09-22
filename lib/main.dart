import 'dart:io';

import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:intl/date_symbol_data_local.dart';
import 'package:intl/intl.dart';
import 'package:shared_preferences/shared_preferences.dart';

import 'ExcelData.dart';

const int hourRowIndex = 2;
const int dateRowIndex = 1;
const int nameColumnIndex = 0;

ValueNotifier<double?> progress = ValueNotifier(null);
const String outFilePath = 'output_file_path';
const String inFilePath = 'input_file_path';

late SharedPreferences prefs;
void main() async {
  WidgetsFlutterBinding.ensureInitialized();
  await initializeDateFormatting('ar');
  prefs = await SharedPreferences.getInstance();
  runApp( MainApp());
}

class MainApp extends StatelessWidget {
  MainApp({super.key});

  late final TextEditingController controller = TextEditingController(text:prefs.getString(outFilePath) ?? '');

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      home: Scaffold(
        body: Padding(
          padding: const EdgeInsets.all(16.0),
          child: Column(
            mainAxisAlignment: MainAxisAlignment.start,
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Text('File Output'),
              Row(
                children: [
                  Expanded(
                    child: TextField(
                      controller:  controller,
                    ),
                  ),
                  IconButton(onPressed: (){
                    FilePicker.platform.pickFiles().then((value){
                      controller.text = value?.files.single.path ?? controller.text;
                      prefs.setString(outFilePath, controller.text);
                    });
                  }, icon: Icon(Icons.file_upload_outlined))
                ],
              ),
              SizedBox(height: 16,),
              Text('select file'),
              MaterialButton(
                color: Colors.blue,
                textColor: Colors.white,

                onPressed: getData,
                child: Text('open'),
              ),
              SizedBox(height: 16),
              Center(
                child: ValueListenableBuilder(
                  valueListenable: progress,
                  builder: (context, value, child) {
                    return Visibility(
                      visible: value != null,
                      child: SizedBox(
                        width: 300,
                        child: Column(
                          children: [
                            LinearProgressIndicator(value: value),
                            Text(
                              "${value == null ? "--" : (value * 100).toStringAsFixed(2)}%",
                            ),
                          ],
                        ),
                      ),
                    );
                  },
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }

  Future<void> getData() async {
    progress.value = null;
    final xFile = await FilePicker.platform.pickFiles();

    Excel excel = Excel.decodeBytes(
      await (xFile?.xFiles.single.readAsBytes() ??
          File('./table.xlsx').readAsBytes()),
    );
    final sheet = excel.tables.values.first;
    progress.value = 0;
    final maxRows = sheet.maxRows;
    final maxColumns = sheet.maxColumns;

    final List<ExcelData> data = [];
    String branchName = "";
    int _progress = 0;
    for (int rowIndex = 5; rowIndex < sheet.maxRows; rowIndex++) {
      final customData = ExcelData(rowIndex);
      DateTime? date;
      progress.value = (_progress++) / maxRows;
      await Future.delayed(const Duration(microseconds: 1));
      for (int columnIndex = 0; columnIndex < sheet.maxColumns; columnIndex++) {
        final data = sheet.cell(
          CellIndex.indexByColumnRow(
            columnIndex: columnIndex,
            rowIndex: rowIndex,
          ),
        );

        final dateCell = sheet.cell(
          CellIndex.indexByColumnRow(
            columnIndex: columnIndex,
            rowIndex: dateRowIndex,
          ),
        );

        final hourCell = sheet.cell(
          CellIndex.indexByColumnRow(
            columnIndex: columnIndex,
            rowIndex: hourRowIndex,
          ),
        );

        if (dateCell.asDateTime != null) {
          date = dateCell.asDateTime;
        }

        if (data.asText.startsWith('فرع')) {
          branchName = data.asText;
          continue;
        } else if (columnIndex == nameColumnIndex) {
          customData.name = data.asText;
          continue;
        }

        if (branchName.isEmpty || (customData.name?.isEmpty ?? true)) {
          continue;
        }

        if (hourCell.asInt != null && data.asInt != null) {
          customData.branchName = branchName;
          customData.quantityByHourAtDateList[date] ??= {};
          customData.quantityByHourAtDateList[date]![hourCell.asInt!] =
              data.asInt!;
        }
      }

      data.add(customData);
    }

    SaveExcelData(data).saveExcel();
  }
}

extension on Data {
  String get asText {
    String val = '';
    if (!isEmpty) {
      val = value.toString().trim();
    }
    return val;
  }

  int? get asInt {
    int? val;
    if (value is IntCellValue?) {
      val = (value as IntCellValue?)?.value;
    }
    return val;
  }

  bool get isEmpty => value == null || value.toString().isEmpty;

  DateTime? get asDateTime {
    DateTime? val;
    if (value is DateTimeCellValue) {
      val = (value as DateTimeCellValue).asDateTimeLocal();
    } else {
      val = _getDate(asText);
    }
    return val;
  }

  DateTime? _getDate(String date) {
    try {
      return DateFormat('dd MMMM yyyy', 'ar').tryParse(date);
    } catch (e) {
      return null;
    }
  }
}
