import 'package:flutter/foundation.dart';

class RosterDoctor {
  const RosterDoctor({
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

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is RosterDoctor &&
        other.id == id &&
        other.name == name &&
        other.departmentId == departmentId &&
        other.teamId == teamId &&
        other.linkedUserId == linkedUserId &&
        other.linkedUserEmail == linkedUserEmail &&
        listEquals(other.positionIds, positionIds);
  }

  @override
  int get hashCode => Object.hash(
    id,
    name,
    departmentId,
    teamId,
    Object.hashAll(positionIds),
    linkedUserId,
    linkedUserEmail,
  );

  @override
  String toString() =>
      'RosterDoctor(id: $id, name: $name, departmentId: $departmentId, teamId: $teamId, positionIds: $positionIds, linkedUserId: $linkedUserId, linkedUserEmail: $linkedUserEmail)';
}
