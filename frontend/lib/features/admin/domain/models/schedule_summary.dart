class ScheduleSummary {
  const ScheduleSummary({
    required this.id,
    required this.departmentId,
    required this.targetMonth,
    required this.sourceFilename,
    required this.totalDuties,
    required this.solverStatus,
    required this.status,
    required this.createdAt,
    required this.updatedAt,
  });

  final String id;
  final String departmentId;
  final String targetMonth;
  final String sourceFilename;
  final int totalDuties;
  final String solverStatus;
  final String status;
  final DateTime createdAt;
  final DateTime updatedAt;

  bool get isPublished => status == 'published';

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is ScheduleSummary &&
        other.id == id &&
        other.departmentId == departmentId &&
        other.targetMonth == targetMonth &&
        other.sourceFilename == sourceFilename &&
        other.totalDuties == totalDuties &&
        other.solverStatus == solverStatus &&
        other.status == status &&
        other.createdAt == createdAt &&
        other.updatedAt == updatedAt;
  }

  @override
  int get hashCode => Object.hash(
    id,
    departmentId,
    targetMonth,
    sourceFilename,
    totalDuties,
    solverStatus,
    status,
    createdAt,
    updatedAt,
  );

  @override
  String toString() =>
      'ScheduleSummary(id: $id, departmentId: $departmentId, targetMonth: $targetMonth, status: $status, totalDuties: $totalDuties)';
}


