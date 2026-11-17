import 'package:frontend/features/admin/domain/models/roster_position.dart';

class RosterPositionDto {
  const RosterPositionDto({
    required this.id,
    required this.name,
    required this.departmentId,
    required this.dutyDays,
  });

  final String id;
  final String name;
  final String departmentId;
  final List<int> dutyDays;

  factory RosterPositionDto.fromJson(Map<String, dynamic> json) {
    final id = json['id'];
    if (id is! String) {
      throw const FormatException('Missing required "id" in position JSON');
    }

    final name = json['name'];
    if (name is! String) {
      throw const FormatException('Missing required "name" in position JSON');
    }

    final departmentId = json['department_id'];
    if (departmentId is! String) {
      throw const FormatException(
        'Missing required "department_id" in position JSON',
      );
    }

    final rawDutyDays = json['duty_days'];
    if (rawDutyDays is! List || rawDutyDays.any((e) => e is! int)) {
      throw const FormatException(
        'Missing or invalid "duty_days" in position JSON',
      );
    }

    return RosterPositionDto(
      id: id,
      name: name,
      departmentId: departmentId,
      dutyDays: List<int>.from(rawDutyDays),
    );
  }

  RosterPosition toDomain() {
    return RosterPosition(
      id: id,
      name: name,
      departmentId: departmentId,
      dutyDays: dutyDays,
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'name': name,
      'department_id': departmentId,
      'duty_days': dutyDays,
    };
  }
}
