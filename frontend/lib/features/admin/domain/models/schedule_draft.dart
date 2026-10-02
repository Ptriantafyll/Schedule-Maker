class ScheduleAssignment {
  const ScheduleAssignment({
    required this.date,
    required this.dayName,
    required this.doctorName,
    required this.doctorEmail,
    required this.position,
    required this.shift,
  });

  final String date;
  final String dayName;
  final String doctorName;
  final String doctorEmail;
  final String position;
  final String shift;

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is ScheduleAssignment &&
        other.date == date &&
        other.dayName == dayName &&
        other.doctorName == doctorName &&
        other.doctorEmail == doctorEmail &&
        other.position == position &&
        other.shift == shift;
  }

  @override
  int get hashCode =>
      Object.hash(date, dayName, doctorName, doctorEmail, position, shift);

  @override
  String toString() =>
      'ScheduleAssignment(date: $date, dayName: $dayName, doctorName: $doctorName, doctorEmail: $doctorEmail, position: $position, shift: $shift)';
}

class ScheduleDraft {
  const ScheduleDraft({
    required this.id,
    required this.departmentId,
    required this.targetMonth,
    required this.sourceFilename,
    required this.totalDuties,
    required this.solverStatus,
    required this.status,
    required this.assignments,
    required this.unavailabilities,
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
  final List<ScheduleAssignment> assignments;
  final Map<String, List<int>> unavailabilities;
  final DateTime createdAt;
  final DateTime updatedAt;

  bool get isPublished => status == 'published';

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is ScheduleDraft &&
        other.id == id &&
        other.departmentId == departmentId &&
        other.targetMonth == targetMonth &&
        other.sourceFilename == sourceFilename &&
        other.totalDuties == totalDuties &&
        other.solverStatus == solverStatus &&
        other.status == status &&
        other.assignments == assignments &&
        other.unavailabilities == unavailabilities &&
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
    assignments,
    unavailabilities,
    createdAt,
    updatedAt,
  );

  @override
  String toString() =>
      'ScheduleDraft(id: $id, departmentId: $departmentId, targetMonth: $targetMonth, sourceFilename: $sourceFilename, totalDuties: $totalDuties, solverStatus: $solverStatus, status: $status, assignments: $assignments, unavailabilities: $unavailabilities, createdAt: $createdAt, updatedAt: $updatedAt)';
}
