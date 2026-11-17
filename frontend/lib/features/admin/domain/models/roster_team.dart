class RosterTeam {
  const RosterTeam({
    required this.id,
    required this.name,
    required this.departmentId,
  });

  final String id;
  final String name;
  final String departmentId;

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is RosterTeam &&
        other.id == id &&
        other.name == name &&
        other.departmentId == departmentId;
  }


  @override
  int get hashCode => Object.hash(
    id,
    departmentId,
    name,
  );

  @override
  String toString() =>
      'RosterTeam(id: $id, name: $name departmentId: $departmentId)';

}
