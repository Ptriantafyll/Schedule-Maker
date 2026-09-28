import 'package:flutter/material.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:frontend/features/admin/presentation/screens/admin_screen.dart';
import 'package:frontend/features/auth/domain/models/user.dart';
import 'package:frontend/features/auth/domain/state/auth_state.dart';
import 'package:frontend/features/auth/presentation/controllers/auth_controller.dart';
import 'package:frontend/features/auth/presentation/screens/login_screen.dart';
import 'package:frontend/features/auth/presentation/screens/signup_screen.dart';

class AuthGate extends ConsumerStatefulWidget {
  const AuthGate({super.key, this.authenticatedBuilder, this.home});

  final Widget? home;
  final Widget Function(BuildContext context, User user)? authenticatedBuilder;

  @override
  ConsumerState<AuthGate> createState() => _AuthGateState();
}

class _AuthGateState extends ConsumerState<AuthGate> {
  bool _showSignup = false;

  @override
  Widget build(BuildContext context) {
    final authState = ref.watch(authControllerProvider);

    if (authState.isLoading && !authState.hasValue) {
      return const Scaffold(body: Center(child: CircularProgressIndicator()));
    }

    final stateValue = authState.valueOrNull;

    if (stateValue is AuthStateAuthenticated) {
      final user = stateValue.user;
      if (widget.authenticatedBuilder != null) {
        return widget.authenticatedBuilder!(context, user);
      }

      if (widget.home != null) {
        return widget.home!;
      }

      return AdminScreen();
      // return Scaffold(
      //   appBar: AppBar(
      //     title: const Text('MedShift Dashboard'),
      //     actions: [
      //       IconButton(
      //         icon: const Icon(Icons.logout),
      //         tooltip: 'Log Out',
      //         onPressed: () {
      //           ref.read(authControllerProvider.notifier).logout();
      //         },
      //       ),
      //     ],
      //   ),
      //   body: Center(
      //     child: Column(
      //       mainAxisAlignment: MainAxisAlignment.center,
      //       children: [
      //         Text(
      //           'Authenticated: ${user.fullName}',
      //           style: const TextStyle(
      //             fontSize: 18,
      //             fontWeight: FontWeight.bold,
      //           ),
      //         ),
      //         const SizedBox(height: 8),
      //         Text('Role: ${user.role.displayName}'),
      //         const SizedBox(height: 24),
      //         FilledButton.icon(
      //           icon: const Icon(Icons.logout),
      //           label: const Text('Log Out'),
      //           onPressed: () {
      //             ref.read(authControllerProvider.notifier).logout();
      //           },
      //         ),
      //       ],
      //     ),
      //   ),
      // );
    }

    if (_showSignup) {
      return SignupScreen(
        onNavigateToLogin: () => setState(() => _showSignup = false),
        onSignupSuccess: (_) => setState(() => _showSignup = false),
      );
    }

    return LoginScreen(
      onNavigateToSignup: () => setState(() => _showSignup = true),
    );
  }
}
