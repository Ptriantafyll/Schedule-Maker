import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/presentation/widgets/admin_hero_card.dart';

void main() {
  group('AdminHeroCard Widget Tests', () {
    testWidgets('renders month title, status pill, description, and generate button', (tester) async {
      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            body: AdminHeroCard(
              targetMonth: 'November 2023',
              statusLabel: 'Ready to Generate',
              onGeneratePressed: () {},
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      expect(find.text('November 2023'), findsOneWidget);
      expect(find.text('Ready to Generate'), findsOneWidget);
      expect(find.textContaining('~45 seconds'), findsOneWidget);
      expect(find.text('Generate Schedule'), findsOneWidget);
      expect(find.byIcon(Icons.auto_awesome), findsOneWidget);
    });

    testWidgets('tapping Generate Schedule button invokes onGeneratePressed callback', (tester) async {
      var wasPressed = false;

      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            body: AdminHeroCard(
              targetMonth: 'November 2023',
              onGeneratePressed: () {
                wasPressed = true;
              },
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      final buttonFinder = find.widgetWithText(FilledButton, 'Generate Schedule');
      expect(buttonFinder, findsOneWidget);

      await tester.tap(buttonFinder);
      await tester.pumpAndSettle();

      expect(wasPressed, isTrue);
    });

    testWidgets('shows loading indicator and disables button when isGenerating is true', (tester) async {
      var wasPressed = false;

      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            body: AdminHeroCard(
              targetMonth: 'November 2023',
              isGenerating: true,
              onGeneratePressed: () {
                wasPressed = true;
              },
            ),
          ),
        ),
      );
      await tester.pump();

      expect(find.byType(CircularProgressIndicator), findsOneWidget);

      final buttonFinder = find.byType(FilledButton);
      final button = tester.widget<FilledButton>(buttonFinder);
      expect(button.onPressed, isNull);
      expect(wasPressed, isFalse);
    });

    testWidgets('renders Published badge when isPublished is true', (tester) async {
      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            body: AdminHeroCard(
              targetMonth: '2026-11',
              isPublished: true,
              onGeneratePressed: () {},
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      expect(find.text('Published'), findsOneWidget);
    });

    testWidgets('renders Draft badge when isPublished is false', (tester) async {
      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            body: AdminHeroCard(
              targetMonth: '2026-11',
              isPublished: false,
              onGeneratePressed: () {},
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      expect(find.text('Draft'), findsOneWidget);
    });

    testWidgets('renders popup menu and invokes onMonthSelected when an item is selected', (tester) async {
      String? selectedMonth;

      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            body: AdminHeroCard(
              targetMonth: '2026-11',
              availableMonths: const ['2026-11', '2026-10'],
              onMonthSelected: (month) => selectedMonth = month,
              onGeneratePressed: () {},
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      // Verify dropdown indicator exists
      expect(find.byIcon(Icons.arrow_drop_down), findsOneWidget);

      // Open popup menu
      await tester.tap(find.byType(PopupMenuButton<String>));
      await tester.pumpAndSettle();

      // Find and tap previous month in popup
      final itemFinder = find.text('2026-10');
      expect(itemFinder, findsOneWidget);

      await tester.tap(itemFinder);
      await tester.pumpAndSettle();

      expect(selectedMonth, '2026-10');
    });

    testWidgets('does not render popup button when availableMonths is empty', (tester) async {
      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            body: AdminHeroCard(
              targetMonth: '2026-11',
              availableMonths: const [],
              onGeneratePressed: () {},
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      expect(find.byType(PopupMenuButton<String>), findsNothing);
      expect(find.text('2026-11'), findsOneWidget);
    });
  });
}
