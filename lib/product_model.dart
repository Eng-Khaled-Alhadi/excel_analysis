class Product {
  late String name;
  late String id;
  late String category;
  late double parts;

  Product({
    required this.name,
    required this.category,
    required this.parts,
    required this.id,
  });

  factory Product.fromJson(Map<String, dynamic> json, String id) {
    return Product(
      name: json['name'],
      category: json['category'],
      parts: (json['parts'] as num).toDouble(),
      id: id,
    );
  }

  toJson() {
    return {'name': name, 'category': category, 'parts': parts};
  }
}
