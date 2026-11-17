class DoctorPreAssignment {
  const DoctorPreAssignment({
    required this.id,
    required this.doctorId,
    required this.shiftId,
    required this.date,
  });

  final String id;
  final String doctorId;
  final String shiftId;
  final String date;

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is DoctorPreAssignment &&
        other.id == id &&
        other.doctorId == doctorId &&
        other.shiftId == shiftId &&
        other.date == date;
  }

  @override
  int get hashCode => Object.hash(id, doctorId, shiftId, date);

  @override
  String toString() =>
      'DoctorPreAssignment(id: $id, doctorId: $doctorId, shiftId: $shiftId, date: $date)';
}
