import 'dart:async';

import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:frontend/features/auth/data/repositories/auth_repository.dart';
import 'package:frontend/features/auth/domain/models/user.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';
import 'package:frontend/features/auth/domain/state/auth_state.dart';

class AuthController extends AsyncNotifier<AuthState> {
  @override
  Future<AuthState> build() async {
    final authRepository = ref.watch(authRepositoryProvider);
    final user = await authRepository.restoreSession();

    if (user == null) {
      return const AuthStateUnauthenticated();
    }

    return AuthStateAuthenticated(user);
  }

  Future<void> login({required String email, required String password}) async {
    state = const AsyncValue.loading();

    try {
      final user = await ref
          .read(authRepositoryProvider)
          .login(email: email, password: password);
      state = AsyncValue.data(AuthStateAuthenticated(user));
    } catch (e, st) {
      state = AsyncValue.error(e, st);
      rethrow;
    }
  }

  Future<User> signup({
    required String invitationToken,
    required String firstName,
    required String lastName,
    required String email,
    required String password,
  }) async {
    final user = await ref
        .read(authRepositoryProvider)
        .signup(
          email: email,
          password: password,
          firstName: firstName,
          lastName: lastName,
          invitationToken: invitationToken,
        );
    return user;
  }

  Future<void> logout() async {
    state = const AsyncValue.loading();
    try {
      await ref.read(authRepositoryProvider).logout();
      state = AsyncValue.data(AuthStateUnauthenticated());
    } catch (e, st) {
      state = AsyncValue.error(e, st);
      rethrow;
    }
  }
}

final authControllerProvider = AsyncNotifierProvider<AuthController, AuthState>(
  AuthController.new,
);

final currentUserProvider = Provider<User?>((ref) {
  final authState = ref.watch(authControllerProvider).valueOrNull;

  return authState?.userOrNull;
});

final isAuthenticatedProvider = Provider<bool>((ref) {
  final currentUser = ref.watch(currentUserProvider);

  return currentUser != null;
});

final userRoleProvider = Provider<UserRole?>((ref) {
  final currentUser = ref.watch(currentUserProvider);

  return currentUser?.role;
});
