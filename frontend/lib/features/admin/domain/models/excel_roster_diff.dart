class MissingDoctorDiffItem {
  const MissingDoctorDiffItem({
    required this.id,
    required this.name,
    this.teamName,
  });

  final String id;
  final String name;
  final String? teamName;

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is MissingDoctorDiffItem &&
        other.id == id &&
        other.name == name &&
        other.teamName == teamName;
  }

  @override
  int get hashCode => Object.hash(id, name, teamName);

  @override
  String toString() =>
      'MissingDoctorDiffItem(id: $id, name: $name, teamName: $teamName)';
}

class NewDoctorDiffItem {
  const NewDoctorDiffItem({
    required this.name,
    this.teamName,
    this.positions = const [],
  });

  final String name;
  final String? teamName;
  final List<String> positions;

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is NewDoctorDiffItem &&
        other.name == name &&
        other.teamName == teamName &&
        _listEquals(other.positions, positions);
  }

  @override
  int get hashCode => Object.hash(name, teamName, Object.hashAll(positions));

  @override
  String toString() =>
      'NewDoctorDiffItem(name: $name, teamName: $teamName, positions: $positions)';
}

bool _listEquals<T>(List<T>? a, List<T>? b) {
  if (a == null) return b == null;
  if (b == null || a.length != b.length) return false;
  if (identical(a, b)) return true;
  for (int index = 0; index < a.length; index += 1) {
    if (a[index] != b[index]) return false;
  }
  return true;
}

class ExcelRosterDiff {
  const ExcelRosterDiff({
    required this.hasExistingRoster,
    this.missingFromExcel = const [],
    this.newInExcel = const [],
  });

  final bool hasExistingRoster;
  final List<MissingDoctorDiffItem> missingFromExcel;
  final List<NewDoctorDiffItem> newInExcel;

  bool get requiresReconciliation =>
      hasExistingRoster &&
      (missingFromExcel.isNotEmpty || newInExcel.isNotEmpty);

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is ExcelRosterDiff &&
        other.hasExistingRoster == hasExistingRoster &&
        _listEquals(other.missingFromExcel, missingFromExcel) &&
        _listEquals(other.newInExcel, newInExcel);
  }

  @override
  int get hashCode => Object.hash(
        hasExistingRoster,
        Object.hashAll(missingFromExcel),
        Object.hashAll(newInExcel),
      );

  @override
  String toString() =>
      'ExcelRosterDiff(hasExistingRoster: $hasExistingRoster, missingFromExcel: $missingFromExcel, newInExcel: $newInExcel)';
}
