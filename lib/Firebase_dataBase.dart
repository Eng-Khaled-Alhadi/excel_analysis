import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:excel_file/product_model.dart';
import 'package:flutter/foundation.dart';

class FB {
  static final _collection = FirebaseFirestore.instance.collection('products');

  static Stream<List<Product>> getProducts() {
    return _collection.snapshots().map((snapshot) {
      return snapshot.docs.map((doc) {
        return Product.fromJson(doc.data(), doc.id);
      }).toList()..sort((a, b) => a.category.compareTo(b.category));
    });
  }

  static final FB instance = FB._internal();

  FB._internal();

  List<Product> products = [];
  List<String> categories = [];

  Stream<List<Product>> listenToProducts() {
    return FirebaseFirestore.instance.collection('products').snapshots().map((
      snapshot,
    ) {
      final list = snapshot.docs.map((doc) {
        return Product.fromJson(doc.data(), doc.id);
      }).toList();

      list.sort((a, b) => a.category.compareTo(b.category));

      // تحديث الكاش
      products = list;

      categories = list.map((product) => product.category).toSet().toList();

      return list;
    });
  }

  static Future<void> addOrUpdateProduct(Product product) async {
    try {
      // البحث عن منتج بنفس الاسم بشكل متزامن
      final querySnapshot = await _collection
          .where('name', isEqualTo: product.name)
          .get();

      if (querySnapshot.docs.isNotEmpty) {
        // المنتج موجود - نقوم بالتحديث
        final docId = querySnapshot.docs.first.id;
        await _collection
            .doc(docId)
            .set(product.toJson(), SetOptions(merge: true));
        debugPrint('تم تحديث المنتج: ${product.name}');
      } else {
        // المنتج غير موجود - نقوم بالإضافة
        await _collection.add(product.toJson());
        debugPrint('تم إضافة المنتج: ${product.name}');
      }
    } catch (e) {
      debugPrint('خطأ في addOrUpdateProduct: $e');
      rethrow;
    }
  }

  static Future<void> deleteProduct(String? id) async {
    try {
      if (id == null || id.isEmpty) {
        debugPrint('معرف المنتج فارغ');
        return;
      }
      await _collection.doc(id).delete();
      debugPrint('تم حذف المنتج بنجاح');
    } catch (e) {
      debugPrint('خطأ في deleteProduct: $e');
      rethrow;
    }
  }

  // دالة للتحقق من وجود منتج
  static Future<bool> productExists(String productName) async {
    try {
      final querySnapshot = await _collection
          .where('name', isEqualTo: productName)
          .limit(1)
          .get();
      return querySnapshot.docs.isNotEmpty;
    } catch (e) {
      debugPrint('خطأ في productExists: $e');
      return false;
    }
  }

  // دالة للحصول على منتج بالاسم
  static Future<Product?> getProductByName(String productName) async {
    try {
      final querySnapshot = await _collection
          .where('name', isEqualTo: productName)
          .limit(1)
          .get();

      if (querySnapshot.docs.isEmpty) {
        return null;
      }

      return Product.fromJson(
        querySnapshot.docs.first.data(),
        querySnapshot.docs.first.id,
      );
    } catch (e) {
      debugPrint('خطأ في getProductByName: $e');
      return null;
    }
  }

  // دالة لجلب كل المنتجات مرة واحدة (بدون stream)
  static Future<List<Product>> fetchAllProducts() async {
    try {
      final snapshot = await _collection.get();
      return snapshot.docs.map((doc) {
        return Product.fromJson(doc.data(), doc.id);
      }).toList()..sort((a, b) => a.category.compareTo(b.category));
    } catch (e) {
      debugPrint('خطأ في fetchAllProducts: $e');
      return [];
    }
  }
}
