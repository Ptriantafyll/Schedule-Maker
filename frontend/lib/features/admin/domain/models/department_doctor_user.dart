class DepartmentDoctorUser {
  const DepartmentDoctorUser({
    required this.id,
    required this.fullName,
    required this.email,
    required this.doctorId,
  });

  final String id;
  final String fullName;
  final String email;
  final String doctorId;

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is DepartmentDoctorUser &&
        other.id == id &&
        other.fullName == fullName &&
        other.email == email &&
        other.doctorId == doctorId;
  }

  @override
  int get hashCode => Object.hash(id, fullName, email, doctorId);

  @override
  String toString() =>
      'DepartmentDoctorUser(id: $id, fullName: $fullName, email: $email, doctorId: $doctorId)';
}
