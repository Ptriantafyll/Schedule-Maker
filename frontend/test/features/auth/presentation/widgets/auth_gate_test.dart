import 'dart:async';

import 'package:flutter/material.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/features/auth/data/repositories/auth_repository.dart';
import 'package:frontend/features/auth/domain/models/user.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';
import 'package:frontend/features/auth/presentation/controllers/auth_controller.dart';
import 'package:frontend/features/auth/presentation/screens/login_screen.dart';
import 'package:frontend/features/auth/presentation/screens/signup_screen.dart';
import 'package:frontend/features/auth/presentation/widgets/auth_gate.dart';

class FakeAuthRepository implements AuthRepository {
  User? userToReturnOnLogin;
  User? userToReturnOnSignup;
  User? userToReturnOnRestoreSession;

  Completer<User?>? restoreSessionCompleter;
  Completer<User>? loginCompleter;

  Exception? loginException;
  Exception? signupException;
  Exception? restoreSessionException;

  int loginCallCount = 0;
  int signupCallCount = 0;
  int logoutCallCount = 0;
  int restoreSessionCallCount = 0;

  String? capturedLoginEmail;
  String? capturedLoginPassword;

  @override
  Future<User> login({
    required String email,
    required String password,
  }) async {
    loginCallCount++;
    capturedLoginEmail = email;
    capturedLoginPassword = password;

    if (loginCompleter != null) {
      return loginCompleter!.future;
    }
    if (loginException != null) {
      throw loginException!;
    }
    return userToReturnOnLogin ??
        const User(
          id: 'usr-1',
          email: 'doctor@hospital.org',
          fullName: 'Dr. Gregory House',
          role: UserRole.doctor,
        );
  }

  @override
  Future<User> signup({
    required String invitationToken,
    required String firstName,
    required String lastName,
    required String email,
    required String password,
  }) async {
    signupCallCount++;
    if (signupException != null) {
      throw signupException!;
    }
    return userToReturnOnSignup ??
        User(
          id: 'usr-2',
          email: email,
          fullName: '$firstName $lastName',
          role: UserRole.doctor,
        );
  }

  @override
  Future<void> logout() async {
    logoutCallCount++;
  }

  @override
  Future<User?> restoreSession() async {
    restoreSessionCallCount++;
    if (restoreSessionCompleter != null) {
      return restoreSessionCompleter!.future;
    }
    if (restoreSessionException != null) {
      throw restoreSessionException!;
    }
    return userToReturnOnRestoreSession;
  }
}

void main() {
  late FakeAuthRepository fakeRepository;

  const testUser = User(
    id: 'usr-1',
    email: 'doctor@hospital.org',
    fullName: 'Dr. Gregory House',
    role: UserRole.doctor,
    departmentId: 'dept-cardio',
  );

  setUp(() {
    fakeRepository = FakeAuthRepository();
  });

  Widget createWidgetUnderTest({
    Widget Function(BuildContext context, User user)? authenticatedBuilder,
    Widget? home,
  }) {
    return ProviderScope(
      overrides: [
        authRepositoryProvider.overrideWithValue(fakeRepository),
      ],
      child: MaterialApp(
        home: AuthGate(
          authenticatedBuilder: authenticatedBuilder,
          home: home,
        ),
      ),
    );
  }

  group('AuthGate - Session Restoration & Loading State', () {
    testWidgets(
      'renders CircularProgressIndicator while initial session is restoring',
      (tester) async {
        fakeRepository.restoreSessionCompleter = Completer<User?>();

        await tester.pumpWidget(createWidgetUnderTest());

        // Wait for first frame
        await tester.pump();

        expect(find.byType(CircularProgressIndicator), findsOneWidget);
        expect(find.byType(LoginScreen), findsNothing);
        expect(find.byType(SignupScreen), findsNothing);

        // Complete session restoration
        fakeRepository.restoreSessionCompleter!.complete(null);
        await tester.pumpAndSettle();

        expect(find.byType(CircularProgressIndicator), findsNothing);
        expect(find.byType(LoginScreen), findsOneWidget);
      },
    );

    testWidgets(
      'falls back to LoginScreen when session restoration throws an error',
      (tester) async {
        fakeRepository.restoreSessionException = const ApiException(
          type: ApiErrorType.unauthorized,
          message: 'Session expired',
          statusCode: 401,
        );

        await tester.pumpWidget(createWidgetUnderTest());
        await tester.pumpAndSettle();

        expect(find.byType(LoginScreen), findsOneWidget);
        expect(find.byType(CircularProgressIndicator), findsNothing);
      },
    );
  });

  group('AuthGate - Unauthenticated Flow & View Switching', () {
    testWidgets(
      'renders LoginScreen when session resolves to unauthenticated',
      (tester) async {
        fakeRepository.userToReturnOnRestoreSession = null;

        await tester.pumpWidget(createWidgetUnderTest());
        await tester.pumpAndSettle();

        expect(find.byType(LoginScreen), findsOneWidget);
        expect(find.byType(SignupScreen), findsNothing);
      },
    );

    testWidgets(
      'switches to SignupScreen when tapping "Complete Onboarding Registration"',
      (tester) async {
        fakeRepository.userToReturnOnRestoreSession = null;

        await tester.pumpWidget(createWidgetUnderTest());
        await tester.pumpAndSettle();

        expect(find.byType(LoginScreen), findsOneWidget);
        expect(find.byType(SignupScreen), findsNothing);

        final toSignupBtn = find.text('Complete Onboarding Registration');
        await tester.ensureVisible(toSignupBtn);
        await tester.tap(toSignupBtn);
        await tester.pumpAndSettle();

        expect(find.byType(SignupScreen), findsOneWidget);
        expect(find.byType(LoginScreen), findsNothing);
      },
    );

    testWidgets(
      'switches back to LoginScreen from SignupScreen when tapping "Sign in to Dashboard"',
      (tester) async {
        fakeRepository.userToReturnOnRestoreSession = null;

        await tester.pumpWidget(createWidgetUnderTest());
        await tester.pumpAndSettle();

        // Navigate to Signup
        final toSignupBtn = find.text('Complete Onboarding Registration');
        await tester.ensureVisible(toSignupBtn);
        await tester.tap(toSignupBtn);
        await tester.pumpAndSettle();

        expect(find.byType(SignupScreen), findsOneWidget);

        // Tap Sign in to Dashboard to go back
        final toLoginBtn = find.text('Sign in to Dashboard');
        await tester.ensureVisible(toLoginBtn);
        await tester.tap(toLoginBtn);
        await tester.pumpAndSettle();

        expect(find.byType(LoginScreen), findsOneWidget);
        expect(find.byType(SignupScreen), findsNothing);
      },
    );

    testWidgets(
      'switches back to LoginScreen after successful signup',
      (tester) async {
        fakeRepository.userToReturnOnRestoreSession = null;

        await tester.pumpWidget(createWidgetUnderTest());
        await tester.pumpAndSettle();

        // Navigate to Signup
        final toSignupBtn = find.text('Complete Onboarding Registration');
        await tester.ensureVisible(toSignupBtn);
        await tester.tap(toSignupBtn);
        await tester.pumpAndSettle();

        // Fill in signup form
        await tester.enterText(
          find.byKey(const Key('signup_token_field')),
          'valid-invitation-token',
        );
        await tester.enterText(
          find.byKey(const Key('signup_first_name_field')),
          'Gregory',
        );
        await tester.enterText(
          find.byKey(const Key('signup_last_name_field')),
          'House',
        );
        await tester.enterText(
          find.byKey(const Key('signup_email_field')),
          'doctor@hospital.org',
        );
        await tester.enterText(
          find.byKey(const Key('signup_password_field')),
          'Password123!',
        );

        final submitBtn = find.text('Activate Account');
        await tester.ensureVisible(submitBtn);
        await tester.tap(submitBtn);
        await tester.pumpAndSettle();

        // After successful signup, user should be back on LoginScreen
        expect(find.byType(LoginScreen), findsOneWidget);
        expect(find.byType(SignupScreen), findsNothing);
        expect(
          find.text('Account activated successfully! Please sign in.'),
          findsOneWidget,
        );
      },
    );
  });

  group('AuthGate - Authenticated State & Transitions', () {
    testWidgets(
      'renders home widget when already authenticated on launch',
      (tester) async {
        fakeRepository.userToReturnOnRestoreSession = testUser;

        await tester.pumpWidget(
          createWidgetUnderTest(
            home: const Scaffold(
              body: Center(child: Text('Hospital Dashboard Home')),
            ),
          ),
        );
        await tester.pumpAndSettle();

        expect(find.text('Hospital Dashboard Home'), findsOneWidget);
        expect(find.byType(LoginScreen), findsNothing);
        expect(find.byType(SignupScreen), findsNothing);
      },
    );

    testWidgets(
      'renders authenticatedBuilder with user when provided',
      (tester) async {
        fakeRepository.userToReturnOnRestoreSession = testUser;

        await tester.pumpWidget(
          createWidgetUnderTest(
            authenticatedBuilder: (context, user) {
              return Scaffold(
                body: Center(
                  child: Text('Welcome, ${user.fullName} (${user.role.displayName})'),
                ),
              );
            },
          ),
        );
        await tester.pumpAndSettle();

        expect(
          find.text('Welcome, Dr. Gregory House (Doctor)'),
          findsOneWidget,
        );
      },
    );

    testWidgets(
      'renders default placeholder when no home or builder is provided',
      (tester) async {
        fakeRepository.userToReturnOnRestoreSession = testUser;

        await tester.pumpWidget(createWidgetUnderTest());
        await tester.pumpAndSettle();

        expect(find.text('Authenticated: Dr. Gregory House'), findsOneWidget);
      },
    );

    testWidgets(
      'transitions from LoginScreen to authenticated view upon successful login',
      (tester) async {
        fakeRepository.userToReturnOnRestoreSession = null;
        fakeRepository.userToReturnOnLogin = testUser;

        await tester.pumpWidget(
          createWidgetUnderTest(
            home: const Scaffold(
              body: Center(child: Text('Hospital Dashboard Home')),
            ),
          ),
        );
        await tester.pumpAndSettle();

        expect(find.byType(LoginScreen), findsOneWidget);

        // Fill credentials
        await tester.enterText(
          find.byKey(const Key('login_email_field')),
          'doctor@hospital.org',
        );
        await tester.enterText(
          find.byKey(const Key('login_password_field')),
          'Password123!',
        );

        final loginBtn = find.text('Login to Dashboard');
        await tester.ensureVisible(loginBtn);
        await tester.tap(loginBtn);
        await tester.pumpAndSettle();

        expect(find.text('Hospital Dashboard Home'), findsOneWidget);
        expect(find.byType(LoginScreen), findsNothing);
      },
    );

    testWidgets(
      'transitions from authenticated view to LoginScreen upon logout',
      (tester) async {
        fakeRepository.userToReturnOnRestoreSession = testUser;

        await tester.pumpWidget(
          ProviderScope(
            overrides: [
              authRepositoryProvider.overrideWithValue(fakeRepository),
            ],
            child: Consumer(
              builder: (context, ref, child) {
                return MaterialApp(
                  home: AuthGate(
                    authenticatedBuilder: (context, user) {
                      return Scaffold(
                        body: Center(
                          child: ElevatedButton(
                            onPressed: () {
                              ref
                                  .read(authControllerProvider.notifier)
                                  .logout();
                            },
                            child: const Text('Log Out'),
                          ),
                        ),
                      );
                    },
                  ),
                );
              },
            ),
          ),
        );
        await tester.pumpAndSettle();

        expect(find.text('Log Out'), findsOneWidget);
        expect(find.byType(LoginScreen), findsNothing);

        // Tap Log Out
        await tester.tap(find.text('Log Out'));
        await tester.pumpAndSettle();

        expect(fakeRepository.logoutCallCount, equals(1));
        expect(find.byType(LoginScreen), findsOneWidget);
        expect(find.text('Log Out'), findsNothing);
      },
    );
  });
}
