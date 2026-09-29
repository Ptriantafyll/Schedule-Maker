import 'package:flutter/material.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/presentation/screens/admin_screen.dart';
import 'package:frontend/features/auth/domain/models/user.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';
import 'package:frontend/features/auth/presentation/controllers/auth_controller.dart';
import 'package:frontend/features/admin/presentation/widgets/generation_mode_dialog.dart';
import 'package:frontend/shared/widgets/bottom_nav_bar.dart';
import 'package:frontend/shared/widgets/profile_drawer.dart';

void main() {
  const testAdmin = User(
    id: 'adm-101',
    email: 'admin@hospital.org',
    fullName: 'Dr. Gregory House',
    role: UserRole.departmentAdmin,
    departmentId: 'dept-er',
  );

  Widget createWidgetUnderTest({User? user = testAdmin}) {
    return ProviderScope(
      overrides: [
        currentUserProvider.overrideWithValue(user),
      ],
      child: const MaterialApp(
        home: AdminScreen(),
      ),
    );
  }

  group('AdminScreen Scaffold & Shell Tests', () {
    testWidgets('renders AppBar with branding, profile avatar, and notification icon', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());
      await tester.pumpAndSettle();

      expect(find.text('MedShift Admin'), findsOneWidget);
      expect(
        find.descendant(of: find.byType(AppBar), matching: find.byType(CircleAvatar)),
        findsOneWidget,
      );
      expect(
        find.byWidgetPredicate(
          (widget) => widget is Icon && (widget.icon == Icons.notifications_none || widget.icon == Icons.notifications_outlined || widget.icon == Icons.notifications),
        ),
        findsOneWidget,
      );
      expect(find.byType(BottomNavBar), findsOneWidget);
    });

    testWidgets('tapping circular profile avatar opens drawer showing admin full name', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());
      await tester.pumpAndSettle();

      // Initially drawer is closed
      expect(find.byType(ProfileDrawer), findsNothing);

      // Tap profile avatar in AppBar
      final avatarFinder = find.descendant(of: find.byType(AppBar), matching: find.byType(CircleAvatar));
      expect(avatarFinder, findsOneWidget);
      await tester.tap(avatarFinder);
      await tester.pumpAndSettle();

      // Drawer should now be open
      expect(find.byType(ProfileDrawer), findsOneWidget);
      expect(find.text('Dr. Gregory House'), findsOneWidget);
    });

    testWidgets('tapping bottom navigation tabs updates selected tab', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());
      await tester.pumpAndSettle();

      // Tap Schedule tab
      await tester.tap(find.text('Schedule'));
      await tester.pumpAndSettle();

      // Tap Dashboard tab
      await tester.tap(find.text('Dashboard'));
      await tester.pumpAndSettle();

      // Tap Admin tab
      await tester.tap(find.text('Admin'));
      await tester.pumpAndSettle();
    });

    testWidgets('tapping Generate Schedule button opens GenerationModeDialog', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());
      await tester.pumpAndSettle();

      expect(find.byType(GenerationModeDialog), findsNothing);

      await tester.tap(find.text('Generate Schedule'));
      await tester.pumpAndSettle();

      expect(find.byType(GenerationModeDialog), findsOneWidget);
      expect(find.text('Select Generation Mode'), findsOneWidget);
    });
  });
}
