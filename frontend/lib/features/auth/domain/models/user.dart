import 'package:frontend/features/auth/domain/models/user_role.dart';

class User {
  const User({
    required this.id,
    required this.email,
    required this.fullName,
    required this.role,
    this.departmentId,
    this.doctorId,
  });

  final String id;
  final String email;
  final String fullName;
  final UserRole role;
  final String? departmentId;
  final String? doctorId;

  User copyWith({
    String? id,
    String? email,
    String? fullName,
    UserRole? role,
    String? departmentId,
    String? doctorId,
  }) {
    return User(
      id: id ?? this.id,
      email: email ?? this.email,
      fullName: fullName ?? this.fullName,
      role: role ?? this.role,
      departmentId: departmentId ?? this.departmentId,
      doctorId: doctorId ?? this.doctorId,
    );
  }

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is User &&
        other.id == id &&
        other.email == email &&
        other.fullName == fullName &&
        other.role == role &&
        other.departmentId == departmentId &&
        other.doctorId == doctorId;
  }

  @override
  int get hashCode =>
      Object.hash(id, email, fullName, role, departmentId, doctorId);

  @override
  String toString() =>
      'User(id: $id, email: $email, fullName: $fullName, role: $role, departmentId: $departmentId, doctorId: $doctorId)';
}
