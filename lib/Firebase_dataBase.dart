import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:excel_file/product_model.dart';
import 'package:flutter/foundation.dart';

class FB {
  static final _collection = FirebaseFirestore.instance.collection('products');

  // تحسين: استخدام الكاش بدلاً من استدعاء مستمر
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

  // تحسين: إضافة batch operations للإضافة المتعددة
  static Future<void> addOrUpdateProduct(Product product) async {
    try {
      // البحث عن منتج بنفس الاسم بشكل متزامن
      final querySnapshot = await _collection
          .where('name', isEqualTo: product.name)
          .limit(1)
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

  // تحسين: إضافة batch operations
  static Future<void> addOrUpdateProductsBatch(List<Product> products) async {
    try {
      final batch = FirebaseFirestore.instance.batch();
      int operationCount = 0;

      for (final product in products) {
        final querySnapshot = await _collection
            .where('name', isEqualTo: product.name)
            .limit(1)
            .get();

        if (querySnapshot.docs.isNotEmpty) {
          final docId = querySnapshot.docs.first.id;
          batch.set(
            _collection.doc(docId),
            product.toJson(),
            SetOptions(merge: true),
          );
        } else {
          batch.set(_collection.doc(), product.toJson());
        }

        operationCount++;

        // Firestore يسمح ب 500 عملية في الـ batch
        if (operationCount >= 500) {
          await batch.commit();
          operationCount = 0;
        }
      }

      // تنفيذ العمليات المتبقية
      if (operationCount > 0) {
        await batch.commit();
      }

      debugPrint('تم إضافة/تحديث ${products.length} منتج بنجاح');
    } catch (e) {
      debugPrint('خطأ في addOrUpdateProductsBatch: $e');
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

  // دالة للتحقق من وجود منتج - محسّنة
  static Future<bool> productExists(String productName) async {
    try {
      final querySnapshot = await _collection
          .where('name', isEqualTo: productName)
          .limit(1)
          .get(const GetOptions(source: Source.cache));

      if (querySnapshot.docs.isNotEmpty) return true;

      // إذا لم يوجد في الكاش، ابحث في السيرفر
      final serverSnapshot = await _collection
          .where('name', isEqualTo: productName)
          .limit(1)
          .get();

      return serverSnapshot.docs.isNotEmpty;
    } catch (e) {
      debugPrint('خطأ في productExists: $e');
      return false;
    }
  }

  // دالة للحصول على منتج بالاسم - محسّنة
  static Future<Product?> getProductByName(String productName) async {
    try {
      // البحث في الكاش أولاً
      final cachedProduct = instance.products
          .where((p) => p.name == productName)
          .firstOrNull;

      if (cachedProduct != null) {
        return cachedProduct;
      }

      // إذا لم يوجد في الكاش، ابحث في Firebase
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

  // دالة لجلب كل المنتجات مرة واحدة (بدون stream) - محسّنة
  static Future<List<Product>> fetchAllProducts() async {
    try {
      // محاولة الحصول على البيانات من الكاش أولاً
      final cachedSnapshot = await _collection.get(
        const GetOptions(source: Source.cache),
      );

      if (cachedSnapshot.docs.isNotEmpty) {
        debugPrint('تم تحميل ${cachedSnapshot.docs.length} منتج من الكاش');
        return cachedSnapshot.docs.map((doc) {
          return Product.fromJson(doc.data(), doc.id);
        }).toList()..sort((a, b) => a.category.compareTo(b.category));
      }

      // إذا لم يوجد في الكاش، اجلب من السيرفر
      final snapshot = await _collection.get();
      debugPrint('تم تحميل ${snapshot.docs.length} منتج من السيرفر');

      return snapshot.docs.map((doc) {
        return Product.fromJson(doc.data(), doc.id);
      }).toList()..sort((a, b) => a.category.compareTo(b.category));
    } catch (e) {
      debugPrint('خطأ في fetchAllProducts: $e');
      return [];
    }
  }

  // دالة لمسح الكاش
  static Future<void> clearCache() async {
    try {
      await FirebaseFirestore.instance.clearPersistence();
      debugPrint('تم مسح الكاش بنجاح');
    } catch (e) {
      debugPrint('خطأ في clearCache: $e');
    }
  }

  // دالة لإعادة تحميل البيانات من السيرفر
  static Future<List<Product>> reloadProducts() async {
    try {
      final snapshot = await _collection.get(
        const GetOptions(source: Source.server),
      );

      final list = snapshot.docs.map((doc) {
        return Product.fromJson(doc.data(), doc.id);
      }).toList()..sort((a, b) => a.category.compareTo(b.category));

      // تحديث الكاش المحلي
      instance.products = list;
      instance.categories = list.map((p) => p.category).toSet().toList();

      debugPrint('تم إعادة تحميل ${list.length} منتج من السيرفر');
      return list;
    } catch (e) {
      debugPrint('خطأ في reloadProducts: $e');
      return [];
    }
  }
}
