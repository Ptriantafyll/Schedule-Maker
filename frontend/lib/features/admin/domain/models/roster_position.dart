import 'package:flutter/foundation.dart';

class RosterPosition {
  const RosterPosition({
    required this.id,
    required this.name,
    required this.departmentId,
    required this.dutyDays,
  });

  final String id;
  final String name;
  final String departmentId;
  final List<int> dutyDays;

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is RosterPosition &&
        other.id == id &&
        other.name == name &&
        other.departmentId == departmentId &&
        listEquals(other.dutyDays, dutyDays);
  }

  @override
  int get hashCode =>
      Object.hash(id, name, departmentId, Object.hashAll(dutyDays));

  @override
  String toString() =>
      'RosterPosition(id: $id, name: $name, departmentId: $departmentId, dutyDays: $dutyDays)';
}
