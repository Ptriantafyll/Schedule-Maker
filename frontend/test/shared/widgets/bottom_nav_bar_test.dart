import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/shared/widgets/bottom_nav_bar.dart';

void main() {
  group('BottomNavBar Widget Tests', () {
    testWidgets('renders all 4 navigation destinations when showAdminTab is true', (tester) async {
      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            bottomNavigationBar: BottomNavBar(
              currentIndex: 3,
              showAdminTab: true,
              onTabSelected: (_) {},
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      expect(find.text('Dashboard'), findsOneWidget);
      expect(find.text('Schedule'), findsOneWidget);
      expect(find.text('Requests'), findsOneWidget);
      expect(find.text('Admin'), findsOneWidget);

      expect(find.byIcon(Icons.grid_view), findsWidgets);
      expect(find.byIcon(Icons.calendar_month), findsWidgets);
      expect(find.byIcon(Icons.swap_horiz), findsWidgets);
      expect(find.byIcon(Icons.admin_panel_settings), findsWidgets);
    });

    testWidgets('renders only 3 navigation destinations when showAdminTab is false', (tester) async {
      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            bottomNavigationBar: BottomNavBar(
              currentIndex: 0,
              showAdminTab: false,
              onTabSelected: (_) {},
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      expect(find.text('Dashboard'), findsOneWidget);
      expect(find.text('Schedule'), findsOneWidget);
      expect(find.text('Requests'), findsOneWidget);
      expect(find.text('Admin'), findsNothing);
    });

    testWidgets('tapping a tab triggers onTabSelected with correct index', (tester) async {
      int? selectedIndex;

      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            bottomNavigationBar: BottomNavBar(
              currentIndex: 3,
              showAdminTab: true,
              onTabSelected: (index) {
                selectedIndex = index;
              },
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      // Tap Dashboard (index 0)
      await tester.tap(find.text('Dashboard'));
      await tester.pumpAndSettle();
      expect(selectedIndex, equals(0));

      // Tap Schedule (index 1)
      await tester.tap(find.text('Schedule'));
      await tester.pumpAndSettle();
      expect(selectedIndex, equals(1));

      // Tap Requests (index 2)
      await tester.tap(find.text('Requests'));
      await tester.pumpAndSettle();
      expect(selectedIndex, equals(2));

      // Tap Admin (index 3)
      await tester.tap(find.text('Admin'));
      await tester.pumpAndSettle();
      expect(selectedIndex, equals(3));
    });
  });
}
