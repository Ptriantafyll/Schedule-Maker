enum UserRole {
  superAdmin('super_admin'),
  departmentAdmin('department_admin'),
  doctor('doctor'),
  viewer('viewer');

  const UserRole(this.value);

  final String value;

  factory UserRole.fromString(String value) {
    return UserRole.values.firstWhere(
      (role) => role.value == value,
      orElse: () => throw FormatException('Invalid UserRole: "$value"'),
    );
  }

  bool get isAdmin => this == superAdmin || this == departmentAdmin;
  bool get requiresDepartment => this != superAdmin;
  bool get requiresDoctor => this == doctor;
  String get displayName => switch(this) {
    UserRole.superAdmin => 'Super Admin',
    UserRole.departmentAdmin => 'Department Admin',
    UserRole.doctor => 'Doctor',
    UserRole.viewer => 'Viewer',
  };
}
