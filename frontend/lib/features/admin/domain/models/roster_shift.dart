class RosterShift {
  const RosterShift({
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

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is RosterShift &&
        other.id == id &&
        other.name == name &&
        other.positionId == positionId &&
        other.doctorsPerShift == doctorsPerShift &&
        other.grantsDayOff == grantsDayOff;
  }

  @override
  int get hashCode =>
      Object.hash(id, name, positionId, doctorsPerShift, grantsDayOff);

  @override
  String toString() =>
      'RosterShift(id: $id, name: $name, positionId: $positionId, doctorsPerShift: $doctorsPerShift, grantsDayOff: $grantsDayOff)';
}
