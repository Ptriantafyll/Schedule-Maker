import 'package:frontend/features/admin/domain/models/roster_doctor.dart';

class RosterDoctorDto {
  const RosterDoctorDto({
    required this.id,
    required this.name,
    required this.departmentId,
    this.teamId,
    required this.positionIds,
    this.linkedUserId,
    this.linkedUserEmail,
  });

  final String id;
  final String name;
  final String departmentId;
  final String? teamId;
  final List<String> positionIds;
  final String? linkedUserId;
  final String? linkedUserEmail;

  factory RosterDoctorDto.fromJson(Map<String, dynamic> json) {
    final id = json['id'];
    if (id is! String) {
      throw const FormatException('Missing required "id" in doctor JSON');
    }

    final name = json['name'];
    if (name is! String) {
      throw const FormatException('Missing required "name" in doctor JSON');
    }

    final departmentId = json['department_id'];
    if (departmentId is! String) {
      throw const FormatException(
        'Missing required "department_id" in doctor JSON',
      );
    }

    final teamId = json['team_id'];
    if (teamId != null && teamId is! String) {
      throw const FormatException('Invalid "team_id" in doctor JSON');
    }

    final rawPositionIds = json['position_ids'];
    if (rawPositionIds != null &&
        (rawPositionIds is! List || rawPositionIds.any((e) => e is! String))) {
      throw const FormatException('Invalid "position_ids" in doctor JSON');
    }

    final linkedUserId = json['linked_user_id'];
    if (linkedUserId != null && linkedUserId is! String) {
      throw const FormatException('Invalid "linked_user_id" in doctor JSON');
    }

    final linkedUserEmail = json['linked_user_email'];
    if (linkedUserEmail != null && linkedUserEmail is! String) {
      throw const FormatException('Invalid "linked_user_email" in doctor JSON');
    }

    return RosterDoctorDto(
      id: id,
      name: name,
      departmentId: departmentId,
      teamId: teamId as String?,
      positionIds: rawPositionIds == null
          ? const []
          : List<String>.from(rawPositionIds),
      linkedUserId: linkedUserId as String?,
      linkedUserEmail: linkedUserEmail as String?,
    );
  }

  RosterDoctor toDomain() {
    return RosterDoctor(
      id: id,
      name: name,
      departmentId: departmentId,
      teamId: teamId,
      positionIds: positionIds,
      linkedUserId: linkedUserId,
      linkedUserEmail: linkedUserEmail,
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'name': name,
      'department_id': departmentId,
      'team_id': teamId,
      'position_ids': positionIds,
      'linked_user_id': ?linkedUserId,
      'linked_user_email': ?linkedUserEmail,
    };
  }
}
