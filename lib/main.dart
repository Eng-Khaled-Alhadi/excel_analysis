import 'dart:io';
import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:intl/date_symbol_data_local.dart';
import 'package:intl/intl.dart';

Future<void> main() async{
  WidgetsFlutterBinding.ensureInitialized();
  runApp(MyApp());


}

class DataModel {
  String? category;
  int? parts;
  String? name;
  String? branch;
  DateTime? date;
  int? hour;
  int? quantity;
  int? price;

  int get total => (quantity ?? 0) * (parts ?? 1);

}


DateTime? getDate(String date){
  try{
    return DateFormat('dd MMMM yyyy','ar').tryParse(date);
  }catch(e){
    return null;
  }
}


class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,

      home: Home(),
    );
  }
}


class Home extends StatelessWidget {
  const Home({super.key});

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: Column(
        mainAxisAlignment: MainAxisAlignment.center,

        children: [
          MaterialButton(onPressed: (){

          },
          child: Text("اختر ملف"),
          )
        ],
      ),
    );
  }
}
