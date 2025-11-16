import 'dart:io';

import 'package:excel/excel.dart' hide Border;
import 'package:excel_file/Firebase_dataBase.dart';
import 'package:excel_file/firebase_options.dart';
import 'package:excel_file/product_model.dart';
import 'package:file_picker/file_picker.dart';
import 'package:firebase_core/firebase_core.dart';
import 'package:flutter/services.dart';
import 'package:intl/date_symbol_data_local.dart';
import 'package:flutter/material.dart';
import 'package:intl/intl.dart';
import 'package:shared_preferences/shared_preferences.dart';
import 'ExcelData.dart';
import 'dart:ui' as ui;

// ============================================
// Constants
// ============================================
const int hourRowIndex = 2;
const int dateRowIndex = 1;
const int nameColumnIndex = 0;
const String outFilePath = 'output_file_path';
const String inFilePath = 'input_file_path';

// ============================================
// Global Variables
// ============================================
ValueNotifier<double?> progress = ValueNotifier(null);
late SharedPreferences prefs;

// ============================================
// Main Entry Point
// ============================================
void main() async {
  WidgetsFlutterBinding.ensureInitialized();
  await Firebase.initializeApp(options: DefaultFirebaseOptions.currentPlatform);
  await initializeDateFormatting('ar');
  prefs = await SharedPreferences.getInstance();
  runApp(const MainApp());
}

// ============================================
// Main App Widget
// ============================================
class MainApp extends StatelessWidget {
  const MainApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        primaryColor: const Color(0xFFFF5722),
        scaffoldBackgroundColor: const Color(0xFFFFF8F5),
        colorScheme: ColorScheme.fromSeed(
          seedColor: const Color(0xFFFF5722),
          primary: const Color(0xFFFF5722),
          secondary: const Color(0xFFFF8A65),
        ),
        fontFamily: 'Tajawal',
        textTheme: const TextTheme(
          bodyLarge: TextStyle(fontSize: 16),
          bodyMedium: TextStyle(fontSize: 15),
          bodySmall: TextStyle(fontSize: 14),
          labelLarge: TextStyle(fontSize: 16),
        ),
      ),
      home: const HomePage(),
    );
  }
}

// ============================================
// Home Page (Stateful Widget)
// ============================================
class HomePage extends StatefulWidget {
  const HomePage({super.key});

  @override
  State<HomePage> createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {
  // Controllers
  late final TextEditingController _outputPathController;
  final TextEditingController _nameController = TextEditingController();
  final TextEditingController _personsController = TextEditingController();

  // State Variables
  bool _addProductShow = false;
  String? _selectedCategory;
  bool _loading = false;

  @override
  void initState() {
    super.initState();
    _outputPathController = TextEditingController(
      text: prefs.getString(outFilePath) ?? '',
    );
  }

  @override
  void dispose() {
    _outputPathController.dispose();
    _nameController.dispose();
    _personsController.dispose();
    super.dispose();
  }

  // ============================================
  // UI Build Method
  // ============================================
  @override
  Widget build(BuildContext context) {
    return Directionality(
      textDirection: ui.TextDirection.rtl,
      child: Scaffold(
        backgroundColor: const Color(0xFFFFF8F5),
        appBar: AppBar(
          backgroundColor: const Color(0xFFFF5722),
          foregroundColor: Colors.white,
          title: Row(
            mainAxisAlignment: MainAxisAlignment.center,
            mainAxisSize: MainAxisSize.min,
            children: [
              Image.asset(
                'assets/images/Icon.png',
                height: 100,
                errorBuilder: (context, error, stackTrace) {
                  return const Icon(Icons.restaurant, size: 36);
                },
              ),
              const SizedBox(width: 12),
              const Text(
                'إدارة المنتجات والملفات',
                style: TextStyle(fontWeight: FontWeight.bold, fontSize: 20),
              ),
            ],
          ),
          centerTitle: true,
          elevation: 4,
        ),
        body: Padding(
          padding: const EdgeInsets.all(20.0),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              _buildFileSection(),
              const SizedBox(height: 20),
              _buildProgressSection(),
              const SizedBox(height: 20),
              Expanded(child: _buildProductsSection()),
            ],
          ),
        ),
      ),
    );
  }

  // ============================================
  // File Section Widget
  // ============================================
  Widget _buildFileSection() {
    return Card(
      elevation: 3,
      color: Colors.white,
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
      child: Padding(
        padding: const EdgeInsets.all(18.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const Text(
              'إعدادات الملف',
              style: TextStyle(
                fontSize: 20,
                fontWeight: FontWeight.bold,
                color: Color(0xFFFF5722),
              ),
            ),
            const SizedBox(height: 16),
            const Text(
              'الملف المستخرج:',
              style: TextStyle(fontSize: 16, fontWeight: FontWeight.w500),
            ),
            const SizedBox(height: 10),
            Row(
              children: [
                Expanded(
                  child: TextField(
                    controller: _outputPathController,
                    style: const TextStyle(fontSize: 15),
                    decoration: InputDecoration(
                      hintText: 'حدد مسار الملف المستخرج',
                      hintStyle: const TextStyle(fontSize: 15),
                      border: OutlineInputBorder(
                        borderRadius: BorderRadius.circular(8),
                      ),
                      focusedBorder: OutlineInputBorder(
                        borderRadius: BorderRadius.circular(8),
                        borderSide: const BorderSide(
                          color: Color(0xFFFF5722),
                          width: 2,
                        ),
                      ),
                      contentPadding: const EdgeInsets.symmetric(
                        horizontal: 14,
                        vertical: 14,
                      ),
                      filled: true,
                      fillColor: Colors.white,
                    ),
                  ),
                ),
                const SizedBox(width: 10),
                IconButton(
                  onPressed: _selectOutputFile,
                  icon: const Icon(
                    Icons.folder_open,
                    color: Color(0xFFFF5722),
                    size: 28,
                  ),
                  tooltip: 'اختيار المجلد',
                  style: IconButton.styleFrom(
                    backgroundColor: const Color(0xFFFFE0D6),
                    padding: const EdgeInsets.all(14),
                  ),
                ),
              ],
            ),
            const SizedBox(height: 18),
            Center(
              child: ElevatedButton.icon(
                onPressed: _getData,
                icon: const Icon(Icons.upload_file, size: 22),
                label: const Text(
                  'تحديد ملف Excel للمعالجة',
                  style: TextStyle(fontSize: 17),
                ),
                style: ElevatedButton.styleFrom(
                  backgroundColor: const Color(0xFFFF5722),
                  foregroundColor: Colors.white,
                  padding: const EdgeInsets.symmetric(
                    horizontal: 28,
                    vertical: 16,
                  ),
                  shape: RoundedRectangleBorder(
                    borderRadius: BorderRadius.circular(8),
                  ),
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }

  // ============================================
  // Progress Section Widget
  // ============================================
  Widget _buildProgressSection() {
    return ValueListenableBuilder(
      valueListenable: progress,
      builder: (context, value, child) {
        if (value == null) return const SizedBox.shrink();

        return Card(
          elevation: 3,
          color: Colors.white,
          shape: RoundedRectangleBorder(
            borderRadius: BorderRadius.circular(12),
          ),
          child: Padding(
            padding: const EdgeInsets.all(18.0),
            child: Column(
              children: [
                const Text(
                  'جاري معالجة الملف...',
                  style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold),
                ),
                const SizedBox(height: 14),
                ClipRRect(
                  borderRadius: BorderRadius.circular(8),
                  child: LinearProgressIndicator(
                    value: value,
                    minHeight: 12,
                    backgroundColor: const Color(0xFFFFE0D6),
                    valueColor: const AlwaysStoppedAnimation<Color>(
                      Color(0xFFFF5722),
                    ),
                  ),
                ),
                const SizedBox(height: 10),
                Text(
                  "${(value * 100).toStringAsFixed(1)}%",
                  style: const TextStyle(
                    fontSize: 20,
                    fontWeight: FontWeight.bold,
                    color: Color(0xFFFF5722),
                  ),
                ),
              ],
            ),
          ),
        );
      },
    );
  }

  // ============================================
  // Products Section Widget
  // ============================================
  Widget _buildProductsSection() {
    return Card(
      elevation: 3,
      color: Colors.white,
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
      child: Padding(
        padding: const EdgeInsets.all(18.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            Row(
              mainAxisAlignment: MainAxisAlignment.spaceBetween,
              children: [
                const Text(
                  'المنتجات',
                  style: TextStyle(
                    fontSize: 20,
                    fontWeight: FontWeight.bold,
                    color: Color(0xFFFF5722),
                  ),
                ),
                Row(
                  children: [
                    ElevatedButton.icon(
                      onPressed: _importProductsFromExcel,
                      icon: const Icon(Icons.file_download, size: 20),
                      label: const Text(
                        'استيراد المنتجات',
                        style: TextStyle(fontSize: 16),
                      ),
                      style: ElevatedButton.styleFrom(
                        backgroundColor: const Color(0xFF2196F3),
                        foregroundColor: Colors.white,
                        padding: const EdgeInsets.symmetric(
                          horizontal: 20,
                          vertical: 12,
                        ),
                        shape: RoundedRectangleBorder(
                          borderRadius: BorderRadius.circular(8),
                        ),
                      ),
                    ),
                    const SizedBox(width: 10),
                    ElevatedButton.icon(
                      onPressed: _toggleAddProductBox,
                      icon: Icon(
                        _addProductShow ? Icons.close : Icons.add,
                        size: 20,
                      ),
                      label: Text(
                        _addProductShow ? 'إلغاء' : 'إضافة منتج',
                        style: const TextStyle(fontSize: 16),
                      ),
                      style: ElevatedButton.styleFrom(
                        backgroundColor: _addProductShow
                            ? Colors.red[400]
                            : const Color(0xFF4CAF50),
                        foregroundColor: Colors.white,
                        padding: const EdgeInsets.symmetric(
                          horizontal: 20,
                          vertical: 12,
                        ),
                        shape: RoundedRectangleBorder(
                          borderRadius: BorderRadius.circular(8),
                        ),
                      ),
                    ),
                  ],
                ),
              ],
            ),
            const SizedBox(height: 16),
            if (_addProductShow) _buildAddProductForm(),
            if (_addProductShow) const SizedBox(height: 16),
            _loading
                ? Column(
                    children: [
                      CircularProgressIndicator(),
                      SizedBox(width: 10),
                      Text('جاري استيراد المنتجات...'),
                    ],
                  )
                : Expanded(child: _buildProductsTable()),
          ],
        ),
      ),
    );
  }

  // ============================================
  // Add Product Form Widget
  // ============================================
  Widget _buildAddProductForm() {
    return Container(
      padding: const EdgeInsets.all(18),
      decoration: BoxDecoration(
        color: const Color(0xFFFFE0D6),
        borderRadius: BorderRadius.circular(8),
        border: Border.all(color: const Color(0xFFFF8A65)),
      ),
      child: Column(
        children: [
          Row(
            children: [
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    const Text(
                      'اسم المنتج *',
                      style: TextStyle(
                        fontWeight: FontWeight.w600,
                        fontSize: 15,
                      ),
                    ),
                    const SizedBox(height: 8),
                    TextField(
                      controller: _nameController,
                      style: const TextStyle(fontSize: 15),
                      decoration: InputDecoration(
                        hintText: 'اكتب اسم المنتج',
                        hintStyle: const TextStyle(fontSize: 15),
                        border: OutlineInputBorder(
                          borderRadius: BorderRadius.circular(8),
                        ),
                        focusedBorder: OutlineInputBorder(
                          borderRadius: BorderRadius.circular(8),
                          borderSide: const BorderSide(
                            color: Color(0xFFFF5722),
                            width: 2,
                          ),
                        ),
                        filled: true,
                        fillColor: Colors.white,
                        contentPadding: const EdgeInsets.symmetric(
                          horizontal: 14,
                          vertical: 14,
                        ),
                      ),
                    ),
                  ],
                ),
              ),
              const SizedBox(width: 16),
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    const Text(
                      'التصنيف *',
                      style: TextStyle(
                        fontWeight: FontWeight.w600,
                        fontSize: 15,
                      ),
                    ),
                    const SizedBox(height: 8),
                    DropdownButtonFormField<String>(
                      value: _selectedCategory,
                      style: const TextStyle(fontSize: 15, color: Colors.black),
                      decoration: InputDecoration(
                        border: OutlineInputBorder(
                          borderRadius: BorderRadius.circular(8),
                        ),
                        focusedBorder: OutlineInputBorder(
                          borderRadius: BorderRadius.circular(8),
                          borderSide: const BorderSide(
                            color: Color(0xFFFF5722),
                            width: 2,
                          ),
                        ),
                        filled: true,
                        fillColor: Colors.white,
                        contentPadding: const EdgeInsets.symmetric(
                          horizontal: 14,
                          vertical: 14,
                        ),
                      ),
                      hint: const Text(
                        'اختر التصنيف',
                        style: TextStyle(fontSize: 15),
                      ),
                      items: const [
                        DropdownMenuItem(
                          value: 'الدجاج الباربكيو',
                          child: Text('الدجاج الباربكيو'),
                        ),
                        DropdownMenuItem(
                          value: 'الدجاج الشواية',
                          child: Text('الدجاج الشواية'),
                        ),
                        DropdownMenuItem(
                          value: 'الدجاج المندي',
                          child: Text('الدجاج المندي'),
                        ),
                        DropdownMenuItem(
                          value: 'اللحم الزبدة',
                          child: Text('اللحم الزبدة'),
                        ),
                        DropdownMenuItem(
                          value: 'اللحم المندي',
                          child: Text('اللحم المندي'),
                        ),
                      ],
                      onChanged: (value) {
                        setState(() {
                          _selectedCategory = value;
                        });
                      },
                    ),
                  ],
                ),
              ),
              const SizedBox(width: 16),
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    const Text(
                      'النفر *',
                      style: TextStyle(
                        fontWeight: FontWeight.w600,
                        fontSize: 15,
                      ),
                    ),
                    const SizedBox(height: 8),
                    TextField(
                      controller: _personsController,
                      style: const TextStyle(fontSize: 15),
                      decoration: InputDecoration(
                        hintText: 'عدد الأنفار',
                        hintStyle: const TextStyle(fontSize: 15),
                        border: OutlineInputBorder(
                          borderRadius: BorderRadius.circular(8),
                        ),
                        focusedBorder: OutlineInputBorder(
                          borderRadius: BorderRadius.circular(8),
                          borderSide: const BorderSide(
                            color: Color(0xFFFF5722),
                            width: 2,
                          ),
                        ),
                        filled: true,
                        fillColor: Colors.white,
                        contentPadding: const EdgeInsets.symmetric(
                          horizontal: 14,
                          vertical: 14,
                        ),
                      ),
                      keyboardType: const TextInputType.numberWithOptions(
                        decimal: true,
                      ),
                      inputFormatters: [
                        FilteringTextInputFormatter.allow(
                          RegExp(r'^\d*\.?\d*$'),
                        ),
                      ],
                    ),
                  ],
                ),
              ),
            ],
          ),
          const SizedBox(height: 18),
          ElevatedButton.icon(
            onPressed: _onAddProductPressed,
            icon: const Icon(Icons.check, size: 20),
            label: const Text('حفظ المنتج', style: TextStyle(fontSize: 17)),
            style: ElevatedButton.styleFrom(
              backgroundColor: const Color(0xFF4CAF50),
              foregroundColor: Colors.white,
              padding: const EdgeInsets.symmetric(horizontal: 28, vertical: 16),
              shape: RoundedRectangleBorder(
                borderRadius: BorderRadius.circular(8),
              ),
            ),
          ),
        ],
      ),
    );
  }

  // ============================================
  // Products Table Widget
  // ============================================
  Widget _buildProductsTable() {
    return StreamBuilder<List<Product>>(
      stream: FB.getProducts(),
      builder: (context, snapshot) {
        if (snapshot.connectionState == ConnectionState.waiting) {
          return const Center(
            child: Padding(
              padding: EdgeInsets.all(20.0),
              child: CircularProgressIndicator(
                valueColor: AlwaysStoppedAnimation<Color>(Color(0xFFFF5722)),
              ),
            ),
          );
        }

        final products = snapshot.data ?? [];
        if (products.isEmpty) {
          return const Center(
            child: Padding(
              padding: EdgeInsets.all(40.0),
              child: Column(
                mainAxisAlignment: MainAxisAlignment.center,
                children: [
                  Icon(
                    Icons.inventory_2_outlined,
                    size: 80,
                    color: Colors.grey,
                  ),
                  SizedBox(height: 20),
                  Text(
                    'لا توجد منتجات حالياً',
                    style: TextStyle(fontSize: 18, color: Colors.grey),
                  ),
                ],
              ),
            ),
          );
        }

        return Container(
          decoration: BoxDecoration(
            border: Border.all(color: Colors.grey.shade300),
            borderRadius: BorderRadius.circular(8),
          ),
          child: ClipRRect(
            borderRadius: BorderRadius.circular(8),
            child: SingleChildScrollView(
              scrollDirection: Axis.vertical,
              child: SingleChildScrollView(
                scrollDirection: Axis.horizontal,
                child: DataTable(
                  headingRowColor: MaterialStateProperty.all(
                    const Color(0xFFFF5722),
                  ),
                  headingTextStyle: const TextStyle(
                    color: Colors.white,
                    fontWeight: FontWeight.bold,
                    fontSize: 16,
                    fontFamily: 'Tajawal',
                  ),
                  dataRowMinHeight: 55,
                  dataRowMaxHeight: 65,
                  columnSpacing: 50,
                  dataTextStyle: const TextStyle(
                    fontSize: 15,
                    fontFamily: 'Tajawal',
                  ),
                  columns: const [
                    DataColumn(label: Text('م.')),
                    DataColumn(label: Text('اسم المنتج')),
                    DataColumn(label: Text('التصنيف')),
                    DataColumn(label: Text('النفر')),
                    DataColumn(label: Text('تعديل')),
                    DataColumn(label: Text('حذف')),
                  ],
                  rows: products
                      .asMap()
                      .entries
                      .map(
                        (e) => DataRow(
                          color: MaterialStateProperty.resolveWith<Color?>(
                            (states) => e.key.isEven
                                ? const Color(0xFFFFF8F5)
                                : Colors.white,
                          ),
                          cells: [
                            DataCell(
                              Text(
                                (e.key + 1).toString(),
                                style: const TextStyle(
                                  fontWeight: FontWeight.w600,
                                  fontSize: 15,
                                ),
                              ),
                            ),
                            DataCell(
                              Text(
                                e.value.name,
                                style: const TextStyle(
                                  fontWeight: FontWeight.w600,
                                  fontSize: 15,
                                ),
                              ),
                            ),
                            DataCell(
                              Text(
                                e.value.category,
                                style: const TextStyle(fontSize: 15),
                              ),
                            ),
                            DataCell(
                              Text(
                                e.value.parts.toString(),
                                style: const TextStyle(
                                  fontWeight: FontWeight.w600,
                                  fontSize: 15,
                                ),
                              ),
                            ),
                            DataCell(
                              IconButton(
                                onPressed: () {
                                  _toggleAddProductBox(value: true);
                                  _nameController.text = e.value.name;
                                  _selectedCategory = e.value.category;
                                  _personsController.text = e.value.parts
                                      .toString();
                                },
                                icon: const Icon(
                                  Icons.edit,
                                  color: Color(0xFFFF5722),
                                  size: 22,
                                ),
                                tooltip: 'تعديل',
                              ),
                            ),
                            DataCell(
                              IconButton(
                                onPressed: () => _deleteProduct(e.value.id),
                                icon: const Icon(
                                  Icons.delete,
                                  color: Colors.red,
                                  size: 22,
                                ),
                                tooltip: 'حذف',
                              ),
                            ),
                          ],
                        ),
                      )
                      .toList(),
                ),
              ),
            ),
          ),
        );
      },
    );
  }

  // ============================================
  // Helper Methods
  // ============================================

  void _toggleAddProductBox({bool? value}) {
    setState(() {
      _addProductShow = value ?? !_addProductShow;
    });
  }

  Future<void> _selectOutputFile() async {
    final result = await FilePicker.platform.pickFiles();
    if (result != null) {
      setState(() {
        _outputPathController.text = result.files.single.path ?? '';
        prefs.setString(outFilePath, _outputPathController.text);
      });
    }
  }

  Future<void> _getData() async {
    progress.value = null;
    final xFile = await FilePicker.platform.pickFiles();
    if (xFile == null) return;

    try {
      Excel excel = Excel.decodeBytes(await xFile.xFiles.single.readAsBytes());
      final sheet = excel.tables.values.first;
      progress.value = 0;
      final maxRows = sheet.maxRows;

      final List<ExcelData> data = [];
      String branchName = "";
      int progressCounter = 0;

      for (int rowIndex = 5; rowIndex < sheet.maxRows; rowIndex++) {
        final customData = ExcelData(rowIndex);
        DateTime? date;
        progress.value = (progressCounter++) / maxRows;
        await Future.delayed(const Duration(microseconds: 1));

        for (
          int columnIndex = 0;
          columnIndex < sheet.maxColumns;
          columnIndex++
        ) {
          final cellData = sheet.cell(
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

          if (cellData.asText.startsWith('فرع')) {
            branchName = cellData.asText;
            continue;
          } else if (columnIndex == nameColumnIndex) {
            customData.name = cellData.asText;
            continue;
          }

          if (branchName.isEmpty || (customData.name?.isEmpty ?? true)) {
            continue;
          }

          if (hourCell.asInt != null && cellData.asInt != null) {
            customData.branchName = branchName;
            customData.quantityByHourAtDateList[date] ??= {};
            customData.quantityByHourAtDateList[date]![hourCell.asInt!] =
                cellData.asInt!;
          }
        }

        data.add(customData);
      }

      SaveExcelData(data).saveExcel();
      progress.value = null;

      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(
            content: Text('تمت معالجة الملف بنجاح'),
            backgroundColor: Colors.green,
          ),
        );
      }
    } catch (e) {
      progress.value = null;
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(
            content: Text('حدث خطأ أثناء معالجة الملف: $e'),
            backgroundColor: Colors.red,
          ),
        );
      }
    }
  }

  // ============================================
  // Import Products from Excel
  // ============================================
  Future<void> _importProductsFromExcel() async {
    setState(() {
      _loading = true;
    });
    try {
      final result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['xlsx', 'xls'],
      );

      if (result == null) return;

      // قراءة الملف
      final bytes =
          result.files.single.bytes ??
          await File(result.files.single.path!).readAsBytes();

      Excel excel = Excel.decodeBytes(bytes);
      final sheet = excel.tables.values.first;

      int successCount = 0;
      int errorCount = 0;

      // البدء من الصف الثاني (تخطي الهيدر)
      for (int rowIndex = 1; rowIndex < sheet.maxRows; rowIndex++) {
        try {
          // قراءة البيانات من الأعمدة
          final categoryCell = sheet
              .cell(
                CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: rowIndex),
              )
              .value;
          final nameCell = sheet
              .cell(
                CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: rowIndex),
              )
              .value;
          final partsCell = sheet
              .cell(
                CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: rowIndex),
              )
              .value;

          // التحقق من أن البيانات ليست فارغة
          if (categoryCell == null || nameCell == null || partsCell == null) {
            continue;
          }

          final String category = categoryCell.toString().trim();
          final String name = nameCell.toString().trim();
          final double parts = double.tryParse(partsCell.toString()) ?? 0;

          if (category.isEmpty || name.isEmpty || parts <= 0) {
            errorCount++;
            continue;
          }

          // إضافة أو تحديث المنتج
          final product = Product(
            name: name,
            category: category,
            parts: parts,
            id: '',
          );

          await FB.addOrUpdateProduct(product);
          successCount++;
        } catch (e) {
          errorCount++;
          print('خطأ في الصف $rowIndex: $e');
        }
      }

      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(
            content: Text(
              'تم استيراد $successCount منتج بنجاح${errorCount > 0 ? ' وفشل $errorCount' : ''}',
            ),
            backgroundColor: Colors.green,
            duration: const Duration(seconds: 3),
          ),
        );
      }
    } catch (e) {
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(
            content: Text('حدث خطأ أثناء استيراد الملف: $e'),
            backgroundColor: Colors.red,
          ),
        );
      }
      print('خطأ في استيراد الملف: $e');
    }

    setState(() {
      _loading = false;
    });
  }

  void _onAddProductPressed() {
    final name = _nameController.text.trim();
    final category = _selectedCategory;
    final personsText = _personsController.text.trim();

    if (name.isEmpty || category == null || personsText.isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(
          content: Text('الرجاء تعبئة جميع الحقول'),
          backgroundColor: Colors.orange,
        ),
      );
      return;
    }

    final persons = double.tryParse(personsText);
    if (persons == null) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(
          content: Text('قيمة النفر غير صحيحة'),
          backgroundColor: Colors.red,
        ),
      );
      return;
    }

    final product = Product(
      name: name,
      category: category,
      parts: persons,
      id: '',
    );

    FB
        .addOrUpdateProduct(product)
        .then((_) {
          _nameController.clear();
          _personsController.clear();
          setState(() {
            _selectedCategory = null;
            _addProductShow = false;
          });

          ScaffoldMessenger.of(context).showSnackBar(
            const SnackBar(
              content: Text('تمت إضافة/تحديث المنتج بنجاح'),
              backgroundColor: Colors.green,
            ),
          );
        })
        .catchError((error) {
          ScaffoldMessenger.of(context).showSnackBar(
            SnackBar(
              content: Text('حدث خطأ أثناء الإضافة: $error'),
              backgroundColor: Colors.red,
            ),
          );
        });
  }

  void _deleteProduct(String productId) {
    showDialog(
      context: context,
      builder: (context) => AlertDialog(
        title: const Text('تأكيد الحذف'),
        content: const Text('هل أنت متأكد من حذف هذا المنتج؟'),
        actions: [
          TextButton(
            onPressed: () => Navigator.pop(context),
            child: const Text('إلغاء'),
          ),
          TextButton(
            onPressed: () {
              FB.deleteProduct(productId);
              Navigator.pop(context);
              ScaffoldMessenger.of(context).showSnackBar(
                const SnackBar(
                  content: Text('تم حذف المنتج'),
                  backgroundColor: Colors.green,
                ),
              );
            },
            child: const Text('حذف', style: TextStyle(color: Colors.red)),
          ),
        ],
      ),
    );
  }
}

// ============================================
// Extensions
// ============================================
extension DataExtension on Data {
  String get asText {
    if (isEmpty) return '';
    return value.toString().trim();
  }

  int? get asInt {
    if (value is IntCellValue?) {
      return (value as IntCellValue?)?.value;
    }
    return null;
  }

  bool get isEmpty => value == null || value.toString().isEmpty;

  DateTime? get asDateTime {
    if (value is DateTimeCellValue) {
      return (value as DateTimeCellValue).asDateTimeLocal();
    }
    return _getDate(asText);
  }

  DateTime? _getDate(String date) {
    try {
      return DateFormat('dd MMMM yyyy', 'ar').tryParse(date);
    } catch (e) {
      return null;
    }
  }
}
