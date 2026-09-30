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
  });
}
