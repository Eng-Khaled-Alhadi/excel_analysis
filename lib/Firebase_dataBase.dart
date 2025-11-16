import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:excel_file/product_model.dart';

class FB {
  static final _collection = FirebaseFirestore.instance.collection('products');

  static Stream<List<Product>> getProducts() {
    return _collection.snapshots().map((snapshot) {
      return snapshot.docs.map((doc) {
        return Product.fromJson(doc.data(), doc.id);
      }).toList()..sort((a, b) => a.category.compareTo(b.category));
    });
  }

  static Future<void> addOrUpdateProduct(Product product) async {
    try {
      // البحث عن منتج بنفس الاسم
      final querySnapshot = await _collection
          .where('name', isEqualTo: product.name)
          .limit(1)
          .get();

      if (querySnapshot.docs.isNotEmpty) {
        // المنتج موجود - نقوم بالتحديث
        final docId = querySnapshot.docs.first.id;
        await _collection.doc(docId).update(product.toJson());
        print('تم تحديث المنتج: ${product.name}');
      } else {
        // المنتج غير موجود - نقوم بالإضافة
        await _collection.add(product.toJson());
        print('تم إضافة المنتج: ${product.name}');
      }
    } catch (e) {
      print('خطأ في addOrUpdateProduct: $e');
      rethrow;
    }
  }

  static Future<void> deleteProduct(String? id) async {
    if (id == null || id.isEmpty) return;
    await _collection.doc(id).delete();
  }
}
