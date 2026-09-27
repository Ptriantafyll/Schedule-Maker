import 'dart:async';

import 'package:flutter/material.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/features/auth/data/repositories/auth_repository.dart';
import 'package:frontend/features/auth/domain/models/user.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';
import 'package:frontend/features/auth/presentation/screens/signup_screen.dart';

class FakeAuthRepository implements AuthRepository {
  User? userToReturnOnSignup;
  Exception? signupException;
  Completer<User>? signupCompleter;

  int signupCallCount = 0;
  String? capturedInvitationToken;
  String? capturedFirstName;
  String? capturedLastName;
  String? capturedEmail;
  String? capturedPassword;

  @override
  Future<User> login({
    required String email,
    required String password,
  }) {
    throw UnimplementedError();
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
    capturedInvitationToken = invitationToken;
    capturedFirstName = firstName;
    capturedLastName = lastName;
    capturedEmail = email;
    capturedPassword = password;

    if (signupCompleter != null) {
      return signupCompleter!.future;
    }
    if (signupException != null) {
      throw signupException!;
    }
    return userToReturnOnSignup ??
        const User(
          id: 'usr-new',
          email: 'doctor@hospital.org',
          fullName: 'Dr. Gregory House',
          role: UserRole.doctor,
        );
  }

  @override
  Future<void> logout() async {}

  @override
  Future<User?> restoreSession() async => null;
}

void main() {
  late FakeAuthRepository fakeRepository;

  setUp(() {
    fakeRepository = FakeAuthRepository();
  });

  Widget createWidgetUnderTest({
    VoidCallback? onNavigateToLogin,
    void Function(User user)? onSignupSuccess,
    String? initialInvitationToken,
  }) {
    return ProviderScope(
      overrides: [
        authRepositoryProvider.overrideWithValue(fakeRepository),
      ],
      child: MaterialApp(
        home: SignupScreen(
          onNavigateToLogin: onNavigateToLogin,
          onSignupSuccess: onSignupSuccess,
          initialInvitationToken: initialInvitationToken,
        ),
      ),
    );
  }

  group('SignupScreen Widget Tests (Mockup & Presentation)', () {
    testWidgets('renders all MedShift branding and signup form elements', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());
      await tester.pumpAndSettle();

      // Branding Header
      expect(find.text('MedShift'), findsOneWidget);
      expect(find.text('Clinical Scheduling & Management'), findsOneWidget);
      expect(find.byIcon(Icons.add), findsOneWidget);

      // Card Header
      expect(find.text('Complete Registration'), findsOneWidget);

      // Form Input Fields
      expect(find.byKey(const Key('signup_token_field')), findsOneWidget);
      expect(find.byKey(const Key('signup_first_name_field')), findsOneWidget);
      expect(find.byKey(const Key('signup_last_name_field')), findsOneWidget);
      expect(find.byKey(const Key('signup_email_field')), findsOneWidget);
      expect(find.byKey(const Key('signup_password_field')), findsOneWidget);

      // Labels
      expect(find.textContaining(RegExp(r'Invitation Token', caseSensitive: false)), findsOneWidget);
      expect(find.text('First Name'), findsOneWidget);
      expect(find.text('Last Name'), findsOneWidget);
      expect(find.text('Email'), findsOneWidget);
      expect(find.text('Password'), findsOneWidget);

      // Action Button
      expect(find.text('Activate Account'), findsOneWidget);

      // Bottom Navigation & Footer
      expect(find.text('Already have an active account?'), findsOneWidget);
      expect(find.textContaining(RegExp(r'Sign in to Dashboard', caseSensitive: false)), findsOneWidget);
      expect(find.textContaining('Need assistance? Contact the Help Desk (Ext. 4420)'), findsOneWidget);
    });

    testWidgets('populates initial token if passed in constructor', (tester) async {
      await tester.pumpWidget(
        createWidgetUnderTest(initialInvitationToken: 'inv-tok-999'),
      );
      await tester.pumpAndSettle();

      expect(find.text('inv-tok-999'), findsOneWidget);
    });

    testWidgets('validates required fields before calling repository', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());
      await tester.pumpAndSettle();

      final submitFinder = find.text('Activate Account');
      await tester.ensureVisible(submitFinder);
      await tester.tap(submitFinder);
      await tester.pumpAndSettle();

      expect(find.text('Invitation token is required'), findsOneWidget);
      expect(find.text('First name is required'), findsOneWidget);
      expect(find.text('Last name is required'), findsOneWidget);
      expect(find.text('Email is required'), findsOneWidget);
      expect(find.text('Password is required'), findsOneWidget);
      expect(fakeRepository.signupCallCount, equals(0));
    });

    testWidgets('validates email syntax and password length', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());
      await tester.pumpAndSettle();

      await tester.enterText(find.byKey(const Key('signup_token_field')), 'tok-123');
      await tester.enterText(find.byKey(const Key('signup_first_name_field')), 'Gregory');
      await tester.enterText(find.byKey(const Key('signup_last_name_field')), 'House');
      await tester.enterText(find.byKey(const Key('signup_email_field')), 'invalid-email');
      await tester.enterText(find.byKey(const Key('signup_password_field')), 'short');

      final submitFinder = find.text('Activate Account');
      await tester.ensureVisible(submitFinder);
      await tester.tap(submitFinder);
      await tester.pumpAndSettle();

      expect(find.text('Please enter a valid email address'), findsOneWidget);
      expect(find.text('Password must be at least 8 characters'), findsOneWidget);
      expect(fakeRepository.signupCallCount, equals(0));
    });

    testWidgets('toggles password visibility', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());
      await tester.pumpAndSettle();

      final passwordFieldFinder = find.byKey(const Key('signup_password_field'));
      TextField passwordTextField = tester.widget<TextField>(
        find.descendant(of: passwordFieldFinder, matching: find.byType(TextField)),
      );
      expect(passwordTextField.obscureText, isTrue);

      final toggleFinder = find.byKey(const Key('signup_password_toggle'));
      await tester.ensureVisible(toggleFinder);
      await tester.tap(toggleFinder);
      await tester.pumpAndSettle();

      passwordTextField = tester.widget<TextField>(
        find.descendant(of: passwordFieldFinder, matching: find.byType(TextField)),
      );
      expect(passwordTextField.obscureText, isFalse);
    });

    testWidgets('submits valid registration details to repository', (tester) async {
      User? createdUserReceived;

      await tester.pumpWidget(
        createWidgetUnderTest(
          onSignupSuccess: (user) {
            createdUserReceived = user;
          },
        ),
      );
      await tester.pumpAndSettle();

      await tester.enterText(find.byKey(const Key('signup_token_field')), 'token-xyz-123');
      await tester.enterText(find.byKey(const Key('signup_first_name_field')), 'James');
      await tester.enterText(find.byKey(const Key('signup_last_name_field')), 'Wilson');
      await tester.enterText(find.byKey(const Key('signup_email_field')), 'wilson@hospital.org');
      await tester.enterText(find.byKey(const Key('signup_password_field')), 'Secret1234!');

      final submitFinder = find.text('Activate Account');
      await tester.ensureVisible(submitFinder);
      await tester.tap(submitFinder);
      await tester.pumpAndSettle();

      expect(fakeRepository.signupCallCount, equals(1));
      expect(fakeRepository.capturedInvitationToken, equals('token-xyz-123'));
      expect(fakeRepository.capturedFirstName, equals('James'));
      expect(fakeRepository.capturedLastName, equals('Wilson'));
      expect(fakeRepository.capturedEmail, equals('wilson@hospital.org'));
      expect(fakeRepository.capturedPassword, equals('Secret1234!'));

      // Verifies success callback or confirmation
      expect(createdUserReceived, isNotNull);
      expect(find.byType(SnackBar), findsOneWidget);
      expect(find.textContaining('Account activated successfully'), findsOneWidget);
    });

    testWidgets('shows loading indicator during signup submission', (tester) async {
      final completer = Completer<User>();
      fakeRepository.signupCompleter = completer;

      await tester.pumpWidget(createWidgetUnderTest());
      await tester.pumpAndSettle();

      await tester.enterText(find.byKey(const Key('signup_token_field')), 'token-xyz-123');
      await tester.enterText(find.byKey(const Key('signup_first_name_field')), 'James');
      await tester.enterText(find.byKey(const Key('signup_last_name_field')), 'Wilson');
      await tester.enterText(find.byKey(const Key('signup_email_field')), 'wilson@hospital.org');
      await tester.enterText(find.byKey(const Key('signup_password_field')), 'Secret1234!');

      final submitFinder = find.text('Activate Account');
      await tester.ensureVisible(submitFinder);
      await tester.tap(submitFinder);
      await tester.pump();

      expect(find.byType(CircularProgressIndicator), findsOneWidget);

      completer.complete(
        const User(
          id: 'usr-new',
          email: 'wilson@hospital.org',
          fullName: 'Dr. James Wilson',
          role: UserRole.doctor,
        ),
      );
      await tester.pumpAndSettle();

      expect(find.byType(CircularProgressIndicator), findsNothing);
    });

    testWidgets('displays error SnackBar when signup fails', (tester) async {
      fakeRepository.signupException = const ApiException(
        type: ApiErrorType.validation,
        message: 'Invitation has expired or is invalid',
        statusCode: 400,
      );

      await tester.pumpWidget(createWidgetUnderTest());
      await tester.pumpAndSettle();

      await tester.enterText(find.byKey(const Key('signup_token_field')), 'bad-token');
      await tester.enterText(find.byKey(const Key('signup_first_name_field')), 'James');
      await tester.enterText(find.byKey(const Key('signup_last_name_field')), 'Wilson');
      await tester.enterText(find.byKey(const Key('signup_email_field')), 'wilson@hospital.org');
      await tester.enterText(find.byKey(const Key('signup_password_field')), 'Secret1234!');

      final submitFinder = find.text('Activate Account');
      await tester.ensureVisible(submitFinder);
      await tester.tap(submitFinder);
      await tester.pumpAndSettle();

      expect(find.byType(SnackBar), findsOneWidget);
      expect(find.text('Invitation has expired or is invalid'), findsOneWidget);
    });

    testWidgets('tapping Sign In to Dashboard calls navigation callback', (tester) async {
      var loginNavigated = false;

      await tester.pumpWidget(
        createWidgetUnderTest(
          onNavigateToLogin: () {
            loginNavigated = true;
          },
        ),
      );
      await tester.pumpAndSettle();

      final loginFinder = find.textContaining(RegExp(r'Sign in to Dashboard', caseSensitive: false));
      await tester.ensureVisible(loginFinder);
      await tester.tap(loginFinder);
      await tester.pumpAndSettle();

      expect(loginNavigated, isTrue);
    });
  });
}
