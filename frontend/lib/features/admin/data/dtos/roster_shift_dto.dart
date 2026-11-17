import 'package:frontend/features/admin/domain/models/roster_shift.dart';

class RosterShiftDto {
  const RosterShiftDto({
    required this.id,
    required this.name,
    required this.positionId,
    required this.doctorsPerShift,
    required this.grantsDayOff,
  });

  final String id;
  final String name;
  final String positionId;
  final int doctorsPerShift;
  final bool grantsDayOff;

  factory RosterShiftDto.fromJson(Map<String, dynamic> json) {
    final id = json['id'];
    if (id is! String) {
      throw const FormatException('Missing required "id" in shift JSON');
    }

    final name = json['name'];
    if (name is! String) {
      throw const FormatException('Missing required "name" in shift JSON');
    }

    final positionId = json['position_id'];
    if (positionId is! String) {
      throw const FormatException(
        'Missing required "position_id" in shift JSON',
      );
    }

    final doctorsPerShift = json['doctors_per_shift'];
    if (doctorsPerShift is! int) {
      throw const FormatException(
        'Missing required "doctors_per_shift" in shift JSON',
      );
    }

    final grantsDayOff = json['grants_day_off'];
    if (grantsDayOff is! bool) {
      throw const FormatException(
        'Missing required "grants_day_off" in shift JSON',
      );
    }

    return RosterShiftDto(
      id: id,
      name: name,
      positionId: positionId,
      doctorsPerShift: doctorsPerShift,
      grantsDayOff: grantsDayOff,
    );
  }

  RosterShift toDomain() {
    return RosterShift(
      id: id,
      name: name,
      positionId: positionId,
      doctorsPerShift: doctorsPerShift,
      grantsDayOff: grantsDayOff,
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'name': name,
      'position_id': positionId,
      'doctors_per_shift': doctorsPerShift,
      'grants_day_off': grantsDayOff,
    };
  }
}
