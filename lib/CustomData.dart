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