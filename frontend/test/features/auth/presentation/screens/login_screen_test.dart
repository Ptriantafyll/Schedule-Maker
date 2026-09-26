import 'dart:async';

import 'package:flutter/material.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/features/auth/data/repositories/auth_repository.dart';
import 'package:frontend/features/auth/domain/models/user.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';
import 'package:frontend/features/auth/presentation/screens/login_screen.dart';

class FakeAuthRepository implements AuthRepository {
  User? userToReturnOnLogin;
  Exception? loginException;
  Completer<User>? loginCompleter;

  int loginCallCount = 0;
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
  }) {
    throw UnimplementedError();
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

  Widget createWidgetUnderTest({VoidCallback? onNavigateToSignup}) {
    return ProviderScope(
      overrides: [
        authRepositoryProvider.overrideWithValue(fakeRepository),
      ],
      child: MaterialApp(
        home: LoginScreen(
          onNavigateToSignup: onNavigateToSignup,
        ),
      ),
    );
  }

  group('LoginScreen Widget Tests (Mockup & Presentation)', () {
    testWidgets('renders all MedShift branding and layout elements', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());

      // Branding Header
      expect(find.text('MedShift'), findsOneWidget);
      expect(find.text('Clinical Scheduling & Management'), findsOneWidget);
      expect(find.byIcon(Icons.add), findsOneWidget); // Cross badge

      // Card Header
      expect(find.text('Staff Login'), findsOneWidget);

      // Input Fields
      expect(find.byKey(const Key('login_email_field')), findsOneWidget);
      expect(find.byKey(const Key('login_password_field')), findsOneWidget);
      expect(find.text('Email'), findsOneWidget);
      expect(find.text('Password'), findsOneWidget);

      // Buttons and Links
      expect(find.text('Forgot Password?'), findsOneWidget);
      expect(find.text('Login to Dashboard'), findsOneWidget);
      expect(find.text('New staff member?'), findsOneWidget);
      expect(find.textContaining('Complete Onboarding Registration'), findsOneWidget);

      // Footer
      expect(find.textContaining('Need assistance? Contact the Help Desk (Ext. 4420)'), findsOneWidget);
    });

    testWidgets('validates required fields and email syntax on submission', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());

      // Tap submit with empty fields
      await tester.tap(find.text('Login to Dashboard'));
      await tester.pumpAndSettle();

      expect(find.text('Email is required'), findsOneWidget);
      expect(find.text('Password is required'), findsOneWidget);
      expect(fakeRepository.loginCallCount, equals(0));

      // Enter invalid email format
      await tester.enterText(find.byKey(const Key('login_email_field')), 'invalid-email');
      await tester.enterText(find.byKey(const Key('login_password_field')), 'Password123!');
      await tester.tap(find.text('Login to Dashboard'));
      await tester.pumpAndSettle();

      expect(find.text('Please enter a valid email address'), findsOneWidget);
      expect(fakeRepository.loginCallCount, equals(0));
    });

    testWidgets('toggles password visibility when eye icon is clicked', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());

      // Initial state: password obscured
      final passwordFieldFinder = find.byKey(const Key('login_password_field'));
      TextField passwordTextField = tester.widget<TextField>(
        find.descendant(of: passwordFieldFinder, matching: find.byType(TextField)),
      );
      expect(passwordTextField.obscureText, isTrue);

      // Tap visibility toggle icon
      final toggleFinder = find.byKey(const Key('login_password_toggle'));
      expect(toggleFinder, findsOneWidget);
      await tester.tap(toggleFinder);
      await tester.pumpAndSettle();

      // Obscure text should now be false
      passwordTextField = tester.widget<TextField>(
        find.descendant(of: passwordFieldFinder, matching: find.byType(TextField)),
      );
      expect(passwordTextField.obscureText, isFalse);
    });

    testWidgets('submits valid credentials to controller', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());

      await tester.enterText(find.byKey(const Key('login_email_field')), 'doctor@hospital.org');
      await tester.enterText(find.byKey(const Key('login_password_field')), 'Secret123!');

      await tester.tap(find.text('Login to Dashboard'));
      await tester.pumpAndSettle();

      expect(fakeRepository.loginCallCount, equals(1));
      expect(fakeRepository.capturedLoginEmail, equals('doctor@hospital.org'));
      expect(fakeRepository.capturedLoginPassword, equals('Secret123!'));
    });

    testWidgets('shows loading spinner and disables submit button during login', (tester) async {
      final completer = Completer<User>();
      fakeRepository.loginCompleter = completer;

      await tester.pumpWidget(createWidgetUnderTest());

      await tester.enterText(find.byKey(const Key('login_email_field')), 'doctor@hospital.org');
      await tester.enterText(find.byKey(const Key('login_password_field')), 'Secret123!');

      await tester.tap(find.text('Login to Dashboard'));
      await tester.pump(); // Advance one frame to start login

      // Submit button should show progress indicator
      expect(find.byType(CircularProgressIndicator), findsOneWidget);

      // Finish the async call
      completer.complete(
        const User(
          id: 'usr-1',
          email: 'doctor@hospital.org',
          fullName: 'Dr. Gregory House',
          role: UserRole.doctor,
        ),
      );
      await tester.pumpAndSettle();

      expect(find.byType(CircularProgressIndicator), findsNothing);
    });

    testWidgets('displays error SnackBar when login fails', (tester) async {
      fakeRepository.loginException = const ApiException(
        type: ApiErrorType.unauthorized,
        message: 'Invalid email or password',
        statusCode: 401,
      );

      await tester.pumpWidget(createWidgetUnderTest());

      await tester.enterText(find.byKey(const Key('login_email_field')), 'doctor@hospital.org');
      await tester.enterText(find.byKey(const Key('login_password_field')), 'WrongPassword');

      await tester.tap(find.text('Login to Dashboard'));
      await tester.pumpAndSettle();

      expect(find.byType(SnackBar), findsOneWidget);
      expect(find.text('Invalid email or password'), findsOneWidget);
    });

    testWidgets('tapping Complete Onboarding Registration calls navigation callback', (tester) async {
      var signupNavigated = false;

      await tester.pumpWidget(
        createWidgetUnderTest(
          onNavigateToSignup: () {
            signupNavigated = true;
          },
        ),
      );

      await tester.tap(find.textContaining('Complete Onboarding Registration'));
      await tester.pumpAndSettle();

      expect(signupNavigated, isTrue);
    });

    testWidgets('tapping Forgot Password displays IT Help Desk information dialog', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());

      await tester.tap(find.text('Forgot Password?'));
      await tester.pumpAndSettle();

      expect(find.byType(AlertDialog), findsOneWidget);
      expect(find.text('Password Reset'), findsOneWidget);
      expect(find.textContaining('Help Desk'), findsOneWidget);
    });
  });
}
