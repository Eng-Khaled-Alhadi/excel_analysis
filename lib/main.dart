import 'dart:io';

import 'package:excel/excel.dart';
import 'package:excel_file/CustomData.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';

void main() {
  runApp(const MainApp());
}

class MainApp extends StatelessWidget {
  const MainApp({super.key});

  @override
  Widget build(BuildContext context) {
    getData();
    return MaterialApp(
      home: Scaffold(
        body: Center(
          child: Column(
            mainAxisAlignment: MainAxisAlignment.center,
            children: [
              Text('select file'),
              MaterialButton(
                color: Colors.blue,
                textColor: Colors.white,

                onPressed: getData,
              child: Text('open'),

              )
            ],
          ),
        ),
      ),
    );
  }

  getData()async{
    Excel excel = Excel.decodeBytes(File('/Users/eagleimac/Downloads/table (10).xlsx').readAsBytesSync());
    final sheet = excel.tables.values.first;
    final rows = [...sheet.rows];

    final hourRowIndex = 2;
    final dateRowIndex = 1;
    String branchName = '';

    for(int rowIndex =0 ; rowIndex < sheet.maxRows;rowIndex++){
      String r = "";
      for(int columnIndex=0 ;columnIndex < sheet.maxColumns;columnIndex++){
        final data = sheet.cell(CellIndex.indexByColumnRow(columnIndex: columnIndex, rowIndex: rowIndex));
        if (data.asText.startsWith('فرع') ?? false) {
          branchName = data.asText;
          continue;
        }

        if(branchName.isEmpty){
          continue;
        }

        r += "${data.asInt ?? data.asText}, ";

      }
      print(r);
    }
  }
}

extension on Data{

  String get asText{
    String val = '';
    if(!isEmpty ){
      val = value.toString().trim();
    }
    return val;
  }

  int? get asInt{
    int? val;
    if(value is IntCellValue?){
      val = (value as IntCellValue?)?.value;
    }
    return  val;
  }

  bool get isEmpty => value == null || value.toString().isEmpty;
}