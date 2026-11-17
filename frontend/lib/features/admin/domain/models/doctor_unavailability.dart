class DoctorUnavailability {
  const DoctorUnavailability({
    required this.id,
    required this.doctorId,
    required this.date,
  });

  final String id;
  final String doctorId;
  final String date;

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is DoctorUnavailability &&
        other.id == id &&
        other.doctorId == doctorId &&
        other.date == date;
  }

  @override
  int get hashCode => Object.hash(id, doctorId, date);

  @override
  String toString() =>
      'DoctorUnavailability(id: $id, doctorId: $doctorId, date: $date)';
}
