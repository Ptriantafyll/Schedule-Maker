import 'package:frontend/features/auth/domain/models/user.dart';

sealed class AuthState {
  const AuthState();

  bool get isAuthenticated => this is AuthStateAuthenticated;
  User? get userOrNull => switch (this) {
    AuthStateAuthenticated(:final user) => user,
    _ => null,
  };
}

class AuthStateUnauthenticated extends AuthState {
  const AuthStateUnauthenticated();

  @override
  bool operator ==(Object other) {
    return other is AuthStateUnauthenticated;
  }

  @override
  int get hashCode => 0;

  @override
  String toString() => 'AuthStateUnauthenticated()';
}

class AuthStateAuthenticated extends AuthState {
  const AuthStateAuthenticated(this.user);

  final User user;

  @override
  bool operator ==(Object other) {
    return other is AuthStateAuthenticated && other.user == user;
  }

  @override
  int get hashCode => user.hashCode;

  @override
  String toString() => 'AuthStateAuthenticated(user: $user)';
}
