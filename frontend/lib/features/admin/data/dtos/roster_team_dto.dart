import 'package:frontend/features/admin/domain/models/roster_team.dart';

class RosterTeamDto {
  const RosterTeamDto({
    required this.id,
    required this.name,
    required this.departmentId,
  });

  final String id;
  final String name;
  final String departmentId;

  factory RosterTeamDto.fromJson(Map<String, dynamic> json) {
    final id = json['id'];
    if (id is! String) {
      throw const FormatException('Missing required "id" in team JSON');
    }

    final name = json['name'];
    if (name is! String) {
      throw const FormatException('Missing required "name" in team JSON');
    }

    final departmentId = json['department_id'];
    if (departmentId is! String) {
      throw const FormatException(
        'Missing required "department_id" in team JSON',
      );
    }

    return RosterTeamDto(id: id, name: name, departmentId: departmentId);
  }

  RosterTeam toDomain() {
    return RosterTeam(id: id, name: name, departmentId: departmentId);
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'name': name,
      'department_id': departmentId,
    };
  }
}
